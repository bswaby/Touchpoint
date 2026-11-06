# -*- coding: utf-8 -*-
#####################################################################
# TPxi_Texting
#####################################################################
# Texting for PoweredByText (PBT) inside TouchPoint, in one page.
#
# HOME, for staff: the conversations from the PBT groups you are listed in
# on SMS Management. Read, reply, text someone new, or text a whole group,
# without signing in to PBT. (This was TPxi_TextingInbox.)
#
# ADMIN TOOLS, for the Admin role only (checked on the server for every
# admin action, not just hidden on the page):
#   Names        who each texting number belongs to in TouchPoint, and which
#                records need fixing. Changing names in PBT is on hold until
#                PBT allows it (NAME_WRITES).
#   Subscribers  TouchPoint's subscriber list next to PBT's, by cell number.
#   Delivery     when each group text went out and how long it took.
#   Usage        credits, plan used, waste, and what is behind each number.
#   Change log   every name change made in PBT.
# (These were TPxi_PBTNameSync.)
#
# Data saved by the two older scripts keeps its Special Content names
# (TPxi_TextingInbox*, TPxi_PBTNameSync*), so nothing is lost by switching.
# Usage is saved by day (TPxi_PBTNameSyncUsageDays); Refresh on the Usage
# page reads only what is new. If the MorningBatch script still has the
# block TPxi_PBTNameSync added, it calls that old script; remove the block
# when the old script is retired.
#
# The two parts are described in detail at the top of each section below.
#
# OPEN IT AT /PyScriptForm/TPxi_Texting (not /PyScript/).
#####################################################################

import json
import re
import datetime

TEXTING_VERSION = '1.0.1'
# Updates come from the same TPxi script service as the other TPxi tools:
# the page asks DC_API_BASE which version is newest, and an admin's Update
# pulls the new code from DC_API_WORKER and saves it over this script.
DC_SCRIPT_ID = 'TPxi_Texting'
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'
model.Header = ''          # the page draws its own title; TouchPoint's is hidden below

# ================================================================ home: conversations (was TPxi_TextingInbox)
# One screen for staff to read and answer texts sent to the church's
# PoweredByText (PBT) groups, without signing in to PBT.
#
# WHO SEES WHAT (checked
# against the bvcms source)
#   The page needs the Access role (TouchPoint's SMS pages are staff only).
#   Reading : only the groups you are listed in on SMS Management, for
#             everyone. This is stricter than TouchPoint, where ManageSMS
#             reads every group on SMS Messages; that page is still there for
#             oversight. A group you are not listed in is not yours to answer.
#   Replying: the SendSMS role AND membership in that group, as in TouchPoint.
#   New text: only people subscribed to a group you can send from show up.
#   People  : for a group you are listed in, compares TouchPoint's subscribers
#             with PoweredByText's, by cell number. Read only. STOP/START
#             texts reach both; hand edits in PoweredByText reach only it.
#   Group texts sent in PoweredByText's app never reach TouchPoint. When a
#   conversation is opened, the last 7 days of them that this number got are
#   read from PBT's report and shown in place (the page keeps them; nothing
#   is stored).
#   Person links open TouchPoint's own record page, with its own rights.
#
# WHERE THE MESSAGES COME FROM (nothing new is stored for them)
#   Incoming : dbo.SmsReceived (TouchPoint's PBT webhook writes every reply).
#   Outgoing : dbo.SMSList + SMSItems. PBT reports every send back to
#              TouchPoint, including texts sent in PBT itself, and it reports
#              each one up to four times (one row per delivery update), so
#              repeats inside 10 minutes are shown once.
#   A text that went to 3+ numbers from a group on the same day is a
#   broadcast. It shows in the thread for context but never counts as an
#   answer.
#
# SENDING, AND THE TEXTING RULES IT FOLLOWS
#   * To a TouchPoint person: model.SendSms, TouchPoint's own sender. It is
#     logged with the staff member as sender and TouchPoint itself refuses
#     anyone with texts turned off or opted out of that group.
#   * To a number nobody in TouchPoint has: PBT's two-way message call,
#     with autoSubscribe off, so it can only reach someone already
#     subscribed. Who sent it is kept in TPxi_TextingInboxSent.
#   * Opted out (STOP) : the reply box is locked.
#   * New texts        : only to people subscribed to that group in
#                        TouchPoint and not opted out.
#   * Quiet hours      : nothing is sent before 8 a.m. or after 9 p.m.
#                        unless the person texted in within the last 2 hours.
#
# GROUP TEXT: sends one text to everyone subscribed to a group in PBT
# (including people who joined by keyword), through PBT's broadcast call
# with autoSubscribe off. Same rights as a reply (SendSMS + group member),
# quiet hours apply with no exceptions, and the page shows the people count
# and estimated credits before it asks to confirm. Logged in
# TPxi_TextingInboxSent with who sent it.
#
# CREDIT METER: texts are billed in segments. Plain text (the GSM-7 set):
# 160 characters, or 153 per segment when longer. One emoji or curly
# quote/apostrophe/dash switches the whole text to Unicode: 70, then 67 per
# segment. "Make it plain" swaps curly punctuation for straight. The count
# is an estimate; PBT's report has no SMS credit field to check it against.
#
# NEW / ANSWERED / DONE
#   New       : they texted since YOU last opened the conversation. Kept per
#               staff member (TPxi_TextingInboxReads_<PeopleId>), the same
#               model TPxi Go uses on each device.
#   Answered  : a staff text came after their last text. Read from the texts
#               themselves; nothing stored. Not answered is the opposite.
#   Done      : shared; anyone can mark a conversation handled.
#   Reactions ("Liked..."), STOP/HELP and PBT's automatic replies never make
#   a conversation New or count as an answer. Nothing guesses from wording.
#
# WHAT IS STORED, AND WHY IT HOLDS NO PHONE NUMBERS OR TEXT
#   TouchPoint's API serves any Special Content record by name with no
#   sign-in (GET /v1/Contents), and any signed-in user can read any Setting.
#   So the records here never hold a phone number, a name of the person
#   texting, or message text:
#     * A conversation is identified by TouchPoint's own record number for it:
#       "i" + the Id of its first incoming text (SmsReceived), or "o" + the Id
#       of its first sent text (SMSItems) when they have never texted in.
#     * The sent log keeps a SHA-256 fingerprint of each message (to put the
#       sender's name on it later), never the message itself.
#   Staff first names and times are stored. The page refreshes every 15
#   seconds.
#

HOSTS = [
    ("https://api.poweredbytext.com", "https://login.poweredbytext.com/connect/token"),
    ("https://api.ministrybytext.com", "https://login.ministrybytext.com/connect/token"),
]
STATE_CONTENT = 'TPxi_TextingInboxState'
READS_PREFIX = 'TPxi_TextingInboxReads_'
SENT_CONTENT = 'TPxi_TextingInboxSent'
QUIET_START, QUIET_END = 21, 8      # no sends from 9 p.m. until 8 a.m.
QUIET_GRACE_MIN = 120               # unless they texted in this recently
SEND_TITLE = 'Inbox reply'
SKIP_NAME = re.compile(r'^\(Invalid Group\)|^TouchPoint (SYSTEM|OTM)$', re.I)
STOP_WORDS = set(['STOP', 'STOPALL', 'STOP ALL', 'UNSUBSCRIBE', 'CANCEL', 'END', 'QUIT', 'OPTOUT', 'OPT OUT', 'REVOKE'])
START_WORDS = set(['START', 'UNSTOP', 'RESUME', 'SUBSCRIBE', 'OPTIN', 'OPT IN'])
# iPhone and Android reactions to a message ("Liked "...""). Not a question.
REACTION = re.compile(u'^((Liked|Loved|Laughed at|Emphasized|Disliked|Questioned|Le gust\u00f3|Le encant\u00f3|'
                      u'Le divirti\u00f3|Reacted .{1,8} to|Reacted to|Reaccion\u00f3 con .{1,8} a|Removed an? \\S+ from|'
                      u'Quit\u00f3 .{1,20} de)\\s*([\u201c"\u201d]|una? (imagen|foto|video)\\b|an? (image|photo|picture|video|movie|attachment)\\b)'
                      u'|[^\\w\\s]{1,8}\\s+to\\s+[\u201c"\u201d])', re.I | re.U)
# Short acknowledgments that do not call for an answer.
# PBT's automatic texts (auto-response, STOP/HELP/opt-in confirmations). They
# are not a staff answer.
AUTO_OUT = re.compile(r'(to opt[- ]?out|reply STOP|text STOP|Msg ?(&|and) ?data|Reply HELP|'
                      r'subscriptions have been|you have been unsubscribed|not be monitoring)', re.I)
# PBT report message types that are automatic, and ones that are group sends.
AUTO_TYPES = set(['autoresponse', 'stopresponse', 'stopallresponse', 'helpresponse', 'optinresponse',
                  'keywordresponse', 'systemmessage', 'deact', 'notification'])
GROUP_TYPES = set(['broadcast', 'scheduledbroadcast', 'broadcastmessage'])
TS = '%Y-%m-%d %H:%M:%S'
NEW_BASELINE_DAYS = 7               # on first visit, older texts are not New


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


# ---------------------------------------------------------------- helpers
def _param(name, default=''):
    try:
        v = getattr(model.Data, name)
        v = unicode(v).strip()
        return v if v else default
    except:
        return default


def digits10(num):
    d = re.sub(r'\D', '', num or '')
    if len(d) == 11 and d.startswith('1'):
        d = d[1:]
    return d


def norm(s):
    return re.sub(r'\s+', ' ', (s or '').strip()).lower()


def parse_ts(s):
    try:
        return datetime.datetime.strptime((s or '')[:19], TS)
    except Exception:
        return None


def minutes_between(a, b):
    ta, tb = parse_ts(a), parse_ts(b)
    if not ta or not tb:
        return None
    d = tb - ta
    return (d.days * 86400 + d.seconds) / 60.0


def client_now():
    """The staff member's local clock ('YYYY-MM-DD HH:MM:SS'). TouchPoint stores
    texting times in church local time, and the server clock may be in
    another zone, so the browser's clock is the one to compare with."""
    v = _param('now')
    return v if parse_ts(v) else datetime.datetime.now().strftime(TS)


def sql_ints(xs):
    xs = [int(x) for x in xs]
    return ','.join(str(x) for x in xs) if xs else '0'


def sql_strs(xs):
    xs = [x for x in xs if re.match(r'^\d{10}$', x or '')]
    return ','.join("'%s'" % x for x in xs) if xs else "''"


def load_json_content(name, default):
    try:
        raw = model.TextContent(name)
        v = json.loads(raw) if raw else default
        return v if isinstance(v, type(default)) else default
    except Exception:
        return default


def is_stop(body):
    return re.sub(r'[^A-Z ]', '', (body or '').upper()).strip() in STOP_WORDS


def is_start(body):
    return re.sub(r'[^A-Z ]', '', (body or '').upper()).strip() in START_WORDS


def is_reaction(body):
    return bool(REACTION.match((body or '').strip()))


def is_help(body):
    """HELP/INFO: PoweredByText answers these by itself."""
    return re.sub(r'[^A-Z]', '', (body or '').upper()) in ('HELP', 'INFO')


def fingerprint(text):
    """SHA-256 of the message (first 200 characters, spacing and case
    ignored), shortened. Lets the sent log be matched to a text without
    keeping the text."""
    t = norm(text)[:200]
    try:
        import System
        from System.Security.Cryptography import SHA256
        from System.Text import Encoding
        h = SHA256.Create().ComputeHash(Encoding.UTF8.GetBytes(t))
        return System.BitConverter.ToString(h).Replace('-', '').lower()[:20]
    except ImportError:
        import hashlib
        return hashlib.sha256(t.encode('utf-8')).hexdigest()[:20]


NUM_SQL = "RIGHT(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(r.FromNumber,'(',''),')',''),' ',''),'-',''),'+',''),10)"


# ---------------------------------------------------------------- who am I
ME = model.UserPeopleId          # a property, not a method
IS_STAFF = model.UserIsInRole('Access')
HAS_SEND = model.UserIsInRole('SendSMS')


def can_send_in(g):
    return bool(HAS_SEND and g.get('member'))


def send_block_reason(g):
    if not HAS_SEND:
        return 'You can read this group but need the SendSMS role to send texts.'
    if not g.get('member'):
        return ('You can read this group because admins can see every group, but only its members can send from it. '
                'Ask to be added to it on SMS Management.')
    return ''


def admin_sees_all():
    """Admin setting, off by default: admins read every texting group, not
    only the ones they are listed in. Sending still needs membership."""
    try:
        return bool(load_settings().get('admin_all_groups')) and model.UserIsInRole('Admin')
    except Exception:
        return False


def my_name():
    if not ME:
        return ''
    r = q.QuerySqlTop1("SELECT ISNULL(PreferredName, FirstName) First, Name FROM People WHERE PeopleId = %d" % int(ME))
    return r


def my_groups():
    """Groups I can read: only those I am listed in on SMS Management,
    whatever my roles, unless the admin setting lets admins read them all."""
    if not ME:
        return {}
    see_all = admin_sees_all()
    rows = q.QuerySql("""
        SELECT g.ID, g.Name, g.ExternalId, g.ExternalId2,
               CASE WHEN g.ID IN (SELECT gm.GroupID FROM SMSGroupMembers gm
                                  JOIN Users u ON u.UserId = gm.UserID
                                  WHERE u.PeopleId = %d) THEN 1 ELSE 0 END Member
        FROM SMSGroups g
        WHERE g.GroupType = 'PoweredByText' AND g.IsDeleted = 0
          AND ISNULL(g.isDisabled, 0) = 0 AND ISNULL(g.OneTimeCodeFlag, 0) = 0
          AND ISNULL(g.SystemFlag, 0) = 0
        ORDER BY g.Name""" % int(ME))
    out = {}
    for r in rows:
        if SKIP_NAME.search(r.Name or '') or not (r.Member or see_all):
            continue
        out[r.ID] = {'id': r.ID, 'name': r.Name, 'groupId': r.ExternalId, 'orgunitId': r.ExternalId2,
                     'member': bool(r.Member)}
    return out


# ---------------------------------------------------------------- shared state
def load_state():
    return load_json_content(STATE_CONTENT, {})


def save_state(st, now):
    # Drop entries untouched for 120 days so the blob stays small, and any
    # left from before conversations were keyed by record number.
    keep = {}
    for k, v in st.items():
        if not re.match(r'^[io]\d+$', k):
            continue
        last = v.get('done') or ''
        m = minutes_between(last, now)
        if m is None or m < 120 * 1440:
            keep[k] = v
    model.WriteContentText(STATE_CONTENT, safe_json(keep), "")


def load_reads():
    """{conversation id: when I last opened it} for the signed-in staff member."""
    return load_json_content(READS_PREFIX + str(int(ME)), {}) if ME else {}


def save_reads(reads, now):
    keep = {}
    for k, t in reads.items():
        m = minutes_between(t, now)
        if k == '_since' or (re.match(r'^[io]\d+$', k) and (m is None or m < 120 * 1440)):
            keep[k] = t
    model.WriteContentText(READS_PREFIX + str(int(ME)), safe_json(keep), "")


def load_sent():
    """Who sent what from this inbox. Entries hold a fingerprint of the
    message, never the text or the number; older entries that did are
    converted the first time they are read."""
    sent = load_json_content(SENT_CONTENT, [])
    if any('msg' in e or 'num' in e for e in sent):
        clean = []
        for e in sent:
            e = dict(e)
            if 'msg' in e:
                e['b'] = fingerprint(e.pop('msg'))
            e.pop('num', None)
            clean.append(e)
        sent = clean
        try:
            model.WriteContentText(SENT_CONTENT, safe_json(sent), "")
        except Exception:
            pass
    return sent


def log_sent(entry):
    sent = load_sent()
    sent.insert(0, entry)
    model.WriteContentText(SENT_CONTENT, safe_json(sent[:2000]), "")


def thread_ids(pairs):
    """{(gid, num): id} for conversations: 'i' + Id of the first incoming
    text, else 'o' + Id of the first text sent to the number from the group."""
    pairs = list(set(pairs))
    out = {}
    if not pairs:
        return out
    gids = sql_ints(set(p[0] for p in pairs))
    nums = list(set(p[1] for p in pairs))
    for i in range(0, len(nums), 500):
        chunk = nums[i:i + 500]
        for r in q.QuerySql("""
                SELECT r.ToGroupId G, %s Num, MIN(r.Id) Id FROM SmsReceived r WITH (NOLOCK)
                WHERE r.ToGroupId IN (%s) AND %s IN (%s)
                GROUP BY r.ToGroupId, %s""" % (NUM_SQL, gids, NUM_SQL, sql_strs(chunk), NUM_SQL)):
            out[(r.G, r.Num)] = 'i%d' % r.Id
    missing = [p for p in pairs if p not in out]
    if missing:
        for r in q.QuerySql("""
                SELECT l.SendGroupID G, i.Number Num, MIN(i.Id) Id
                FROM SMSItems i WITH (NOLOCK) JOIN SMSList l WITH (NOLOCK) ON l.Id = i.ListID
                WHERE l.SendGroupID IN (%s) AND i.Number IN (%s)
                GROUP BY l.SendGroupID, i.Number""" % (sql_ints(set(p[0] for p in missing)),
                                                       sql_strs(set(p[1] for p in missing)))):
            if (r.G, r.Num) in missing:
                out[(r.G, r.Num)] = 'o%d' % r.Id
    return out


# ---------------------------------------------------------------- people on a number
def people_on(nums, gids):
    """{num: [person]} with texting flags for the given groups."""
    out = {}
    if not nums:
        return out
    rows = q.QuerySql("""
        SELECT p.PeopleId, p.CellPhone, ISNULL(p.PreferredName, p.FirstName) First, p.LastName,
               p.Age, p.PositionInFamilyId Pos, p.FamilyId, f.HeadOfHouseholdId Head,
               ISNULL(p.ReceiveSMS, 0) ReceiveSMS
        FROM People p WITH (NOLOCK)
        JOIN Families f ON f.FamilyId = p.FamilyId
        WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0 AND p.CellPhone IN (%s)""" % sql_strs(nums))
    pids = []
    for r in rows:
        pids.append(r.PeopleId)
        out.setdefault(r.CellPhone, []).append({
            'pid': r.PeopleId, 'name': ((r.First or '') + ' ' + (r.LastName or '')).strip(),
            'first': r.First or '', 'age': r.Age, 'fam': r.FamilyId, 'head': r.Head,
            'adult': r.Pos != 30 and (r.Age is None or r.Age >= 18),
            'receive': bool(r.ReceiveSMS), 'subs': {}, })
    if pids and gids:
        subs = {}
        for s in q.QuerySql("""
                SELECT ToPeopleId, FromGroup, ISNULL(IsSubscribed, 0) Sub, ISNULL(IsOptedOut, 0) Opt
                FROM SMSGroupSubscriptions WITH (NOLOCK)
                WHERE ToPeopleId IN (%s) AND FromGroup IN (%s)""" % (sql_ints(set(pids)), sql_ints(gids))):
            subs[(s.ToPeopleId, s.FromGroup)] = (bool(s.Sub), bool(s.Opt))
        for lst in out.values():
            for p in lst:
                for g in gids:
                    if (p['pid'], g) in subs:
                        p['subs'][g] = subs[(p['pid'], g)]
    return out


def owner(cands, prefer_pid=None):
    """Who to show the thread as: the person TouchPoint matched the last text
    to, else an adult (head of household first), else anyone."""
    if not cands:
        return None
    for c in cands:
        if c['pid'] == prefer_pid:
            return c
    adults = [c for c in cands if c['adult']]
    heads = [c for c in adults if c['pid'] == c['head']]
    return (heads or adults or cands)[0]


# ---------------------------------------------------------------- messages
def broadcasts(gids, since):
    """{(gid, day, message[:200])} sent to 3+ numbers: announcements, not answers."""
    out = set()
    for r in q.QuerySql("""
            SELECT l.SendGroupID G, CONVERT(varchar(10), l.Created, 120) D, LEFT(l.Message, 200) M
            FROM SMSList l WITH (NOLOCK) JOIN SMSItems i WITH (NOLOCK) ON i.ListID = l.Id
            WHERE l.SendGroupID IN (%s) AND l.Created >= '%s'
            GROUP BY l.SendGroupID, CONVERT(varchar(10), l.Created, 120), LEFT(l.Message, 200)
            HAVING COUNT(DISTINCT i.Number) >= 3""" % (sql_ints(gids), since)):
        out.add((r.G, r.D, norm(r.M)))
    return out


def inbound_rows(where):
    return q.QuerySql("""
        SELECT TOP 5000 r.Id, CONVERT(varchar(19), r.DateReceived, 120) T, r.ToGroupId G,
               r.FromPeopleId Pid, r.Body, r.Action, %s Num
        FROM SmsReceived r WITH (NOLOCK)
        WHERE %s
        ORDER BY r.DateReceived DESC""" % (NUM_SQL, where))


def outbound_rows(where):
    return q.QuerySql("""
        SELECT TOP 5000 l.Id, CONVERT(varchar(19), l.Created, 120) T, l.SendGroupID G, l.SenderID Sender,
               l.Title, l.Message, l.ReplyToId, i.Number Num, i.PeopleID Pid, i.Sent,
               i.ResultStatus Status, ISNULL(i.NoOptIn, 0) NoOptIn
        FROM SMSList l WITH (NOLOCK) JOIN SMSItems i WITH (NOLOCK) ON i.ListID = l.Id
        WHERE %s
        ORDER BY l.Created DESC""" % where)


def collapse(events):
    """PBT reports each text several times. Keep one per direction + body
    within 10 minutes (2 for incoming), with the best delivery status."""
    events.sort(key=lambda e: e['t'])
    out = []
    for e in events:
        prev = None
        for o in reversed(out[-12:]):
            if o['dir'] == e['dir'] and norm(o['body']) == norm(e['body']):
                prev = o
                break
        win = 10 if e['dir'] == 'out' else 2
        m = minutes_between(prev['t'], e['t']) if prev else None
        if prev is not None and m is not None and m <= win:
            if e['dir'] == 'out':
                if e.get('delivered') and not prev.get('delivered'):
                    prev['delivered'], prev['status'] = True, e.get('status')
                if e.get('senderPid') and not prev.get('staff'):
                    prev['senderPid'], prev['staff'] = e['senderPid'], e.get('staff')
            continue
        out.append(e)
    return out


def events_for(g, nums_filter, since, bcast):
    """All events (in + out) for the given groups/numbers, collapsed, grouped by (gid, num)."""
    gids = sql_ints(g)
    numsql = (' AND i.Number IN (%s)' % sql_strs(nums_filter)) if nums_filter is not None else ''
    inb = inbound_rows("r.DateReceived >= '%s' AND r.ToGroupId IN (%s)%s" % (
        since, gids, (' AND %s IN (%s)' % (NUM_SQL, sql_strs(nums_filter))) if nums_filter is not None else ''))
    by = {}
    for r in inb:
        by.setdefault((r.G, r.Num), []).append({
            'dir': 'in', 't': r.T, 'body': r.Body or '', 'pid': r.Pid, 'id': r.Id,
            'stop': is_stop(r.Body) or (r.Action or '').upper() in ('STOP', 'STOPALL'),
            'start': is_start(r.Body), 'reaction': is_reaction(r.Body), 'help': is_help(r.Body)})
    if nums_filter is None:
        nums_filter = list(set(k[1] for k in by.keys()))
        # Threads started from this inbox have no incoming text yet.
        for r in q.QuerySql("""
                SELECT DISTINCT l.SendGroupID G, i.Number Num FROM SMSList l WITH (NOLOCK)
                JOIN SMSItems i WITH (NOLOCK) ON i.ListID = l.Id
                WHERE l.Title = '%s' AND l.Created >= '%s' AND l.SendGroupID IN (%s)""" % (SEND_TITLE, since, gids)):
            if r.Num and (r.G, r.Num) not in by:
                by[(r.G, r.Num)] = []
                nums_filter.append(r.Num)
        numsql = ' AND i.Number IN (%s)' % sql_strs(nums_filter)
    senders = {}
    # Texts sent from this inbox straight through PBT (group texts, and
    # replies to numbers nobody in TouchPoint has). TouchPoint stamps them
    # late and without a sender; the log has the real time and the name.
    sent_log = [e for e in load_sent() if e.get('b') and e.get('gid') in g]
    outs = outbound_rows("l.SendGroupID IN (%s) AND l.Created >= '%s'%s" % (gids, since, numsql))
    for r in outs:
        key = (r.G, r.Num)
        if key not in by:
            continue
        b = (r.G, (r.T or '')[:10], norm((r.Message or '')[:200]))
        from_pbt = (r.Title or '') == 'Text Message'
        e = {'dir': 'out', 't': r.T, 'body': r.Message or '', 'id': r.Id,
             'broadcast': b in bcast, 'status': r.Status or '', 'noOptIn': bool(r.NoOptIn),
             'delivered': (r.Status or '').lower().startswith('deliver') or bool(r.Sent),
             'senderPid': None if from_pbt else r.Sender, 'pbt': from_pbt,
             'auto': bool(AUTO_OUT.search(r.Message or ''))}
        # A group text sent from this inbox: TouchPoint only learns of it when
        # PBT's delivery report arrives, minutes later, so its time is late.
        # The inbox's own log has the real time and the sender.
        if from_pbt and sent_log:
            fp = fingerprint(e['body'])
            for gs in sent_log:
                m = minutes_between(gs['t'], e['t'])
                if gs['gid'] == r.G and gs['b'] == fp and m is not None and -5 <= m <= 30:
                    e['staff'], e['t'], e['timeFixed'] = gs.get('by'), gs['t'], True
                    if gs.get('group'):
                        e['broadcast'] = True
                    break
        if e['senderPid']:
            senders[e['senderPid']] = None
        by[key].append(e)
    if senders:
        for r in q.QuerySql("SELECT PeopleId, ISNULL(PreferredName, FirstName) + ' ' + LastName N FROM People WHERE PeopleId IN (%s)"
                            % sql_ints(senders.keys())):
            senders[r.PeopleId] = r.N
    for key in by:
        for e in by[key]:
            if e['dir'] == 'out' and not e.get('staff') and not e.get('broadcast'):
                e['staff'] = senders.get(e.get('senderPid')) if e.get('senderPid') else None
        by[key] = collapse(by[key])
    return by


def thread_facts(gid, evs, cands):
    """Summary used by the list, the thread header and the send check.
    lastAttn is their last text that a person wrote to us: not a reaction,
    not STOP/START/HELP."""
    last_in = None
    last_attn = None
    last_out = None
    opted = False
    for e in evs:
        if e['dir'] == 'in':
            last_in = e
            if e['stop']:
                opted = True
            elif e['start']:
                opted = False
            elif not e['reaction'] and not e.get('help'):
                last_attn = e
        elif not e.get('broadcast') and not e.get('auto'):
            last_out = e
    prefer = last_in['pid'] if last_in else None
    who = owner(cands, prefer)
    # TouchPoint's own flags: texts off on the record, or opted out of this group.
    textable = [c for c in cands if c['receive'] and not c['subs'].get(gid, (False, False))[1]]
    if cands and not textable and not opted:
        why = 'Texts are turned off, or opted out of this group, on every TouchPoint record with this number'
    else:
        why = ''
    target = None
    if textable:
        target = owner(textable, prefer)
    return {'lastIn': last_in, 'lastAttn': last_attn, 'lastOut': last_out, 'opted': opted,
            'who': who, 'target': target, 'blocked': why}


# ---------------------------------------------------------------- PBT (numbers nobody in TouchPoint has)
_conn = {}


def pbt_token():
    if _conn.get('token'):
        return _conn['token']
    cid = model.Setting("PBTClientID", "")
    secret = model.Setting("PBTClientSecret", "")
    if not cid or not secret:
        raise Exception("PBTClientID or PBTClientSecret is not set in Settings.")
    body = "client_id=%s&client_secret=%s&grant_type=client_credentials" % (
        model.UrlEncode(cid), model.UrlEncode(secret))
    errors = []
    for api, token_url in HOSTS:
        raw = model.RestPost(token_url, {"Content-Type": "application/x-www-form-urlencoded"}, body)
        try:
            tok = json.loads(raw).get('access_token')
        except Exception:
            tok = None
        if tok:
            _conn['token'], _conn['api'] = tok, api
            return tok
        errors.append("%s: %s" % (token_url.split('/')[2], (raw or '(empty reply)')[:120]))
    raise Exception("PBT sign-in failed. " + " | ".join(errors))


def pbt_send(g, num, msg, media=None, fallback=None):
    tok = pbt_token()
    body = {'groupId': [g['groupId']], 'message': msg, 'messageType': 'Twoway', 'channelType': 'SMS',
            'subscriberMobileNo': ['1' + num], 'autoSubscribe': False}
    add_media(body, media, fallback)
    raw = model.RestPost(_conn['api'] + '/orgunit/%s/messages' % g['orgunitId'],
                         {"Authorization": "Bearer " + tok, "Accept": "application/json",
                          "Content-Type": "application/json"}, safe_json(body))
    try:
        res = json.loads(raw) if raw and raw.strip() else {}
    except Exception:
        raise Exception("PoweredByText reply was not understood: %s" % (raw or '(empty)')[:200])
    if res.get('isError'):
        raise Exception(res.get('message') or res.get('appCode') or 'PoweredByText refused the text')


# ---------------------------------------------------------------- pictures, links, emoji
# PBT takes these files for a picture message (MMS). US carriers cap an MMS
# at 950 KB; the page shrinks photos to fit before uploading.
MEDIA_KINDS = {'image/jpeg': 'picture', 'image/png': 'picture', 'image/gif': 'picture',
               'application/pdf': 'pdf', 'video/mp4': 'av', 'audio/mpeg': 'av', 'audio/mp3': 'av'}
MEDIA_ALLOWED = {'off': (), 'pictures': ('picture',), 'pictures_pdf': ('picture', 'pdf'),
                 'all': ('picture', 'pdf', 'av')}
MEDIA_MAX = 950 * 1024
# Carriers block texts with public link shorteners, and PBT warns that
# sending them can get the account suspended. Links PBT or TouchPoint made
# are already short and safe.
PUBLIC_SHORTENER = re.compile(r'https?://(www\.)?(bit\.ly|tinyurl\.com|linktr\.ee|linkr\.ee|t\.co|ow\.ly|is\.gd|'
                              r'rebrand\.ly|shorturl\.at|cutt\.ly|tiny\.cc|buff\.ly|t\.ly|rb\.gy)/', re.I)
ALREADY_SHORT = re.compile(r'https?://([\w-]+\.)*(gopbt\.me|tpsdb\.co)/', re.I)
# A host without a dot is one of TouchPoint's special links (rsvplink,
# registerlink...), filled in per person when TouchPoint sends. Never touch.
URL_RE = re.compile(r'https?://[\w-]+\.[^\s<>"]+')
EMOJI_RE = re.compile(u'[\ud800-\udbff][\udc00-\udfff]|[\u2600-\u27bf\u2b00-\u2bff\ufe0f\u200d]')


def has_emoji(text):
    return bool(EMOJI_RE.search(text or u''))


def check_emoji(text, is_group):
    if send_policy()['emoji'] == 'none' and has_emoji(text):
        raise Exception('Emoji are turned off for texts here. Take them out and send again.')


def shorten_links(g, msg):
    """Long links become PBT short links for this group (the same link gives
    back the same short link). Returns (text, note). If PBT will not shorten
    (its M-Click setting is off), the link goes out as typed."""
    if not msg:
        return msg, ''
    if PUBLIC_SHORTENER.search(msg):
        raise Exception('This text has a public short link (like bit.ly). Carriers block those and PoweredByText '
                        'warns it can get the account suspended. Paste the full link instead; it will be shortened safely.')
    notes = []

    def one(m):
        whole = m.group(0)
        url = whole.rstrip('.,;:!?)\'"')
        trail = whole[len(url):]
        if ALREADY_SHORT.match(url) or len(url) <= 32:
            return whole
        try:
            res = pbt_post('/orgunit/%s/groups/%s/shortUrls' % (g['orgunitId'], g['groupId']), {'longUrl': url})
            short = ((res or {}).get('data') or {}).get('url') if not (res or {}).get('isError') else None
        except Exception:
            short = None
        if not short:
            notes.append('PoweredByText did not shorten a link (its M-Click setting may be off), so it went out as typed.')
            return whole
        return short + trail
    return URL_RE.sub(one, msg), ' '.join(sorted(set(notes)))


def parse_media(raw):
    """The file details the page got back from upload_media, checked."""
    if not raw:
        return None
    try:
        media = json.loads(raw)
    except Exception:
        raise Exception('The attachment details were not understood. Attach it again.')
    if not isinstance(media, list) or not media or not all(isinstance(x, dict) and x.get('url') for x in media):
        raise Exception('The attachment details were not understood. Attach it again.')
    kinds = MEDIA_ALLOWED[send_policy()['attach']]
    for x in media:
        if MEDIA_KINDS.get((x.get('mimetype') or '').lower()) not in kinds:
            raise Exception('That kind of attachment is turned off here.')
    return media[:1]


def add_media(body, media, fallback):
    """Turns a text body into a picture message with its plain-text fallback."""
    if not media:
        return
    body['channelType'] = 'MMS'
    body['mediaDetails'] = {'mediaPath': media}
    body['fallbackMessage'] = {'message': fallback or body.get('message') or '', 'channelType': 'SMS'}


def media_fallback(g, text, media, keep_link):
    """The plain text someone gets when their phone cannot show the file."""
    fb, note = shorten_links(g, text)
    if keep_link and media:
        link, _ = shorten_links(g, media[0]['url'])
        kind = MEDIA_KINDS.get((media[0].get('mimetype') or '').lower())
        fb = (fb + '\n' if fb else '') + {'picture': 'Picture', 'pdf': 'File', 'av': 'Media'}.get(kind, 'File') + ': ' + link
    return fb


def pbt_upload(g, name, mime, data_b64):
    """Sends one file to PBT's media upload (multipart form). Returns PBT's
    mediaPath list, which a send then refers to."""
    import System
    from System.Text import Encoding
    from System.IO import StreamReader
    if (mime or '').lower() not in MEDIA_KINDS:
        raise Exception('PoweredByText takes JPG, PNG, GIF, PDF, MP4 and MP3 files.')
    raw = System.Convert.FromBase64String(data_b64)
    if raw.Length > MEDIA_MAX:
        raise Exception('That file is %d KB. A picture message can be at most 950 KB, so send a link instead.'
                        % (raw.Length // 1024))
    tok = pbt_token()
    safe_name = re.sub(r'[^\w.\-]', '_', name or 'file')[:80] or 'file'
    boundary = '----tpxi' + System.Guid.NewGuid().ToString('N')
    head = Encoding.UTF8.GetBytes('--%s\r\nContent-Disposition: form-data; name="file"; filename="%s"\r\n'
                                  'Content-Type: %s\r\n\r\n' % (boundary, safe_name, mime))
    tail = Encoding.UTF8.GetBytes('\r\n--%s--\r\n' % boundary)
    req = System.Net.WebRequest.Create(_conn['api'] + '/orgunit/%s/messages/file' % g['orgunitId'])
    req.Method = 'POST'
    req.ContentType = 'multipart/form-data; boundary=' + boundary
    req.Accept = 'application/json'
    req.Headers.Add('Authorization', 'Bearer ' + tok)
    req.ContentLength = head.Length + raw.Length + tail.Length
    stream = req.GetRequestStream()
    stream.Write(head, 0, head.Length)
    stream.Write(raw, 0, raw.Length)
    stream.Write(tail, 0, tail.Length)
    stream.Close()
    try:
        resp = req.GetResponse()
    except Exception as e:
        resp = getattr(e, 'Response', None) or getattr(getattr(e, 'clsException', None), 'Response', None)
        if resp is None:
            raise Exception('PoweredByText upload failed: %s' % e)
    text = StreamReader(resp.GetResponseStream()).ReadToEnd()
    resp.Close()
    try:
        res = json.loads(text)
    except Exception:
        raise Exception('PoweredByText upload reply was not understood: %s' % (text or '(empty)')[:200])
    paths = ((res.get('data') or {}).get('mediaPath')) if not res.get('isError') else None
    if not paths:
        raise Exception(res.get('message') or res.get('appCode') or 'PoweredByText did not take the file.')
    return paths


def act_upload_media():
    gid = int(_param('gid', '0'))
    g = my_groups().get(gid)
    if not g:
        raise Exception("You don't have rights to that texting group.")
    if send_block_reason(g):
        raise Exception(send_block_reason(g))
    mime = (_param('fmime') or '').lower()
    if MEDIA_KINDS.get(mime) not in MEDIA_ALLOWED[send_policy()['attach']]:
        raise Exception('That kind of attachment is turned off here.')
    paths = pbt_upload(g, _param('fname'), mime, _param('fdata'))
    return {'success': True, 'media': paths}


# ---------------------------------------------------------------- credits
GSM7 = set(u'@\u00a3$\u00a5\u00e8\u00e9\u00f9\u00ec\u00f2\u00c7\n\u00d8\u00f8\r\u00c5\u00e5\u0394_\u03a6\u0393\u039b\u03a9'
           u'\u03a0\u03a8\u03a3\u0398\u039e\u00c6\u00e6\u00df\u00c9 !"#\u00a4%&\'()*+,-./0123456789:;<=>?\u00a1'
           u'ABCDEFGHIJKLMNOPQRSTUVWXYZ\u00c4\u00d6\u00d1\u00dc\u00a7\u00bfabcdefghijklmnopqrstuvwxyz\u00e4\u00f6\u00f1\u00fc\u00e0')
GSM7_EXT = set(u'^{}\\[~]|\u20ac')


def credits(msg):
    """Estimated segments (credits) for one text. Same rules as the page."""
    msg = msg or u''
    units, uni = 0, False
    for ch in msg:
        if ch in GSM7:
            units += 1
        elif ch in GSM7_EXT:
            units += 2
        else:
            uni = True
    n = len(msg) if uni else units    # IronPython strings are UTF-16: an emoji counts 2
    if not n:
        return 0
    first, rest = (70, 67) if uni else (160, 153)
    return 1 if n <= first else -(-n // rest)


def pbt_group_counts(g):
    """{status: count} for the group's PBT subscribers."""
    counts, page = {}, 1
    while page <= 30:
        res = pbt_get_raw('/orgunit/%s/groups/%s/subscribers?pageNumber=%d&pageSize=1000'
                          % (g['orgunitId'], g['groupId'], page))
        if res.get('isError'):
            raise Exception('PoweredByText: %s %s' % (res.get('appCode') or '', res.get('message') or ''))
        data = res.get('data') or {}
        rows = data.get('data') or []
        for r in rows:
            st = (r.get('status') or 'Unknown').strip() or 'Unknown'
            counts[st] = counts.get(st, 0) + 1
        if not rows or page >= (data.get('totalPages') or 1):
            break
        page += 1
    return counts


def active_count(counts):
    return sum(v for k, v in counts.items() if k.lower() == 'active')


def act_group_info():
    gid = int(_param('gid', '0'))
    g = my_groups().get(gid)
    if not g:
        raise Exception("You don't have rights to that texting group.")
    if send_block_reason(g):
        raise Exception(send_block_reason(g))
    counts = pbt_group_counts(g)
    return {'success': True, 'gid': gid, 'active': active_count(counts), 'counts': counts}


def pbt_group_subscribers(g):
    """Every subscriber PoweredByText has for the group (all statuses)."""
    out, page = [], 1
    while page <= 30:
        res = pbt_get_raw('/orgunit/%s/groups/%s/subscribers?pageNumber=%d&pageSize=1000'
                          % (g['orgunitId'], g['groupId'], page))
        if res.get('isError'):
            raise Exception('PoweredByText: %s %s' % (res.get('appCode') or '', res.get('message') or ''))
        data = res.get('data') or {}
        rows = data.get('data') or []
        out.extend(rows)
        if not rows or page >= (data.get('totalPages') or 1):
            break
        page += 1
    return out


# What each mismatch means, worst first. 'okon'/'okoff' rows agree.
PEOPLE_KINDS = [
    ('optedgets', 'Opted out in TouchPoint, getting texts in PoweredByText',
     'TouchPoint says they opted out of this group, but PoweredByText still sends them group texts.'),
    ('pbtknown', 'Only in PoweredByText, person in TouchPoint',
     'They get group texts and TouchPoint has a person on this number, but does not list them as subscribed. They joined inside PoweredByText (keyword, sign-up link or a hand add), which TouchPoint never hears about.'),
    ('pbtunknown', 'Only in PoweredByText, no one in TouchPoint',
     'They get group texts, but no one in TouchPoint has this cell number. It may be a guest, or someone whose TouchPoint record has an old or missing cell.'),
    ('tponly', 'Only in TouchPoint',
     'TouchPoint lists them as subscribed, but PoweredByText does not have them, so they do not get group texts.'),
    ('tpnotpbt', 'Subscribed in TouchPoint, not active in PoweredByText',
     'TouchPoint lists them as subscribed, but PoweredByText is not sending to them (usually they texted STOP there).'),
    ('okon', 'Getting texts', 'Subscribed in TouchPoint and active in PoweredByText. Both agree.'),
    ('okoff', 'Not getting texts',
     'Opted out or inactive, and both agree. Most texted STOP; only they can rejoin, by texting START.'),
]


def act_people():
    """Read-only comparison of one group's subscribers in TouchPoint and in
    PoweredByText, matched by cell number. Nothing is sent or changed."""
    gid = int(_param('gid', '0'))
    g = my_groups().get(gid)
    if not g:
        raise Exception("You don't have rights to that texting group.")

    pbt = {}
    for s in pbt_group_subscribers(g):
        num = digits10(s.get('mobileNo'))
        if len(num) != 10:
            continue
        st = (s.get('status') or 'Unknown').strip() or 'Unknown'
        prev = pbt.get(num)
        if prev and prev['status'].lower() == 'active':
            continue
        pbt[num] = {'status': st, 'name': ((s.get('firstName') or '') + ' ' + (s.get('lastName') or '')).strip()}

    tp = {}
    for r in q.QuerySql("""
            SELECT p.PeopleId, p.CellPhone, ISNULL(p.PreferredName, p.FirstName) First, p.LastName, p.Age,
                   ISNULL(p.ReceiveSMS, 0) ReceiveSMS, ISNULL(s.IsSubscribed, 0) Sub, ISNULL(s.IsOptedOut, 0) Opt
            FROM SMSGroupSubscriptions s WITH (NOLOCK)
            JOIN People p WITH (NOLOCK) ON p.PeopleId = s.ToPeopleId
            WHERE s.FromGroup = %d AND p.IsDeceased = 0 AND p.ArchivedFlag = 0""" % gid):
        num = digits10(r.CellPhone)
        if len(num) != 10 or not (r.Sub or r.Opt):
            continue
        tp.setdefault(num, []).append({
            'pid': r.PeopleId, 'name': ((r.First or '') + ' ' + (r.LastName or '')).strip(), 'age': r.Age,
            'receive': bool(r.ReceiveSMS), 'state': 'opted' if r.Opt else 'sub'})

    # Names for numbers only PoweredByText has, so staff can see who they are.
    extra = people_on([n for n in pbt if n not in tp], [])

    rows, counts = [], {}
    for num in set(pbt) | set(tp):
        p, people = pbt.get(num), tp.get(num) or []
        active = bool(p) and p['status'].lower() == 'active'
        tstate = ('sub' if any(x['state'] == 'sub' for x in people) else 'opted') if people else None
        if active and tstate == 'opted':
            kind = 'optedgets'
        elif active and not tstate:
            kind = 'pbtknown' if extra.get(num) else 'pbtunknown'
        elif tstate == 'sub' and not p:
            kind = 'tponly'
        elif tstate == 'sub' and not active:
            kind = 'tpnotpbt'
        else:
            kind = 'okon' if active else 'okoff'
        counts[kind] = counts.get(kind, 0) + 1
        others = [{'pid': x['pid'], 'name': x['name'], 'age': x['age'], 'receive': x['receive'], 'state': ''}
                  for x in extra.get(num, [])] if not people else []
        rows.append({'num': num, 'kind': kind, 'pbt': p['status'] if p else '', 'pbtName': p['name'] if p else '',
                     'tp': tstate or '', 'people': people or others})
    order = dict((k[0], i) for i, k in enumerate(PEOPLE_KINDS))
    rows.sort(key=lambda r: (order[r['kind']], (r['people'][0]['name'] if r['people'] else r['pbtName'] or 'zzz').lower()))
    return {'success': True, 'gid': gid, 'group': g['name'], 'rows': rows, 'counts': counts,
            'kinds': [{'key': k, 'label': l, 'help': h} for k, l, h in PEOPLE_KINDS]}


def act_group_send():
    gid = int(_param('gid', '0'))
    msg = _param('msg')
    now = client_now()
    g = my_groups().get(gid)
    if not g:
        raise Exception("You don't have rights to that texting group.")
    if send_block_reason(g):
        raise Exception(send_block_reason(g))
    if not msg:
        raise Exception('The message is empty.')
    if len(msg) > 1600:
        raise Exception('Texts can be at most 1,600 characters.')
    hour = parse_ts(now).hour
    if not (QUIET_END <= hour < QUIET_START):
        raise Exception('Group texts cannot go out during quiet hours (9 p.m. to 8 a.m.).')
    check_emoji(msg, True)
    media = parse_media(_param('media'))
    msg, note = shorten_links(g, msg)
    tok = pbt_token()
    body = {'groupId': [g['groupId']], 'message': msg, 'messageType': 'BroadcastMessage',
            'channelType': 'SMS', 'autoSubscribe': False}
    if media:
        add_media(body, media, media_fallback(g, _param('fb') or msg, media, _param('fblink') != '0'))
    raw = model.RestPost(_conn['api'] + '/orgunit/%s/messages' % g['orgunitId'],
                         {"Authorization": "Bearer " + tok, "Accept": "application/json",
                          "Content-Type": "application/json"}, safe_json(body))
    try:
        res = json.loads(raw) if raw and raw.strip() else {}
    except Exception:
        raise Exception("PoweredByText reply was not understood: %s" % (raw or '(empty)')[:200])
    if res.get('isError'):
        raise Exception(res.get('message') or res.get('appCode') or 'PoweredByText refused the group text')
    me = my_name()
    people = int(_param('expected', '0')) if _param('expected', '0').isdigit() else 0
    log_sent({'t': now, 'gid': gid, 'group': True, 'b': fingerprint(msg), 'pid': ME,
              'by': (me.Name if me else ''), 'people': people, 'credits': 3 if media else credits(msg), 'mms': bool(media)})
    return {'success': True, 'note': note}


# ---------------------------------------------------------------- actions
def act_threads():
    groups = my_groups()
    if not groups:
        return {'success': True, 'threads': [], 'groups': []}
    now = client_now()
    days = int(_param('days', '90')) if _param('days', '90').isdigit() else 90
    since = (parse_ts(now) - datetime.timedelta(days=min(days, 400))).strftime(TS)
    gids = list(groups.keys())
    by = events_for(gids, None, since, broadcasts(gids, since))
    people = people_on(list(set(k[1] for k in by)), gids)
    st = load_state()
    reads = load_reads()
    if ME and not reads.get('_since'):
        # First visit: only texts from the last 7 days count as New, so a
        # new user does not start with months of old texts marked New.
        reads['_since'] = (parse_ts(now) - datetime.timedelta(days=NEW_BASELINE_DAYS)).strftime(TS)
        save_reads(reads, now)
    tids = thread_ids([k for k, v in by.items() if v])
    out = []
    for (gid, num), evs in by.items():
        if not evs:
            continue
        cands = people.get(num, [])
        f = thread_facts(gid, evs, cands)
        tid = tids.get((gid, num), '')
        s = st.get(tid, {}) if tid else {}
        mine = reads.get(tid, '') if tid else ''
        # The list shows the last message that matters: not an automatic reply
        # and not a reaction ("Liked ...").
        shown = [e for e in evs if not e.get('auto') and not e.get('reaction')]
        last = shown[-1] if shown else evs[-1]
        attn = f['lastAttn']
        done = s.get('done') or ''
        if f['opted']:
            status = 'optedout'
        elif done and (not attn or done >= attn['t']):
            status = 'done'
        elif f['lastOut'] and (not attn or f['lastOut']['t'] >= attn['t']):
            status = 'answered'
        elif attn:
            status = 'notanswered'
        else:
            status = 'other'
        seen_at = max(mine, done, reads.get('_since') or '')
        new_count = sum(1 for e in evs if e['dir'] == 'in' and not e['reaction'] and not e.get('help')
                        and not e['stop'] and not e['start'] and e['t'] > seen_at)
        out.append({
            'key': '%s|%s' % (gid, num), 'tid': tid, 'gid': gid, 'group': groups[gid]['name'], 'num': num,
            'name': f['who']['name'] if f['who'] else '', 'pid': f['who']['pid'] if f['who'] else None,
            'others': len(cands) - 1 if cands else 0,
            't': last['t'], 'preview': (last['body'] or '')[:120],
            'lastDir': 'auto' if last.get('auto') else 'bc' if last.get('broadcast') else last['dir'],
            'status': status, 'isNew': new_count > 0 and not f['opted'], 'newCount': new_count,
        })
    news = sorted([x for x in out if x['isNew']], key=lambda x: x['t'], reverse=True)
    rest = sorted([x for x in out if not x['isNew']], key=lambda x: x['t'], reverse=True)
    return {'success': True, 'threads': news + rest,
            'groups': [{'id': g['id'], 'name': g['name'], 'canSend': can_send_in(g)}
                       for g in sorted(groups.values(), key=lambda g: g['name'])]}


def thread_context(gid, num, now):
    groups = my_groups()
    if gid not in groups:
        raise Exception("You don't have rights to that texting group.")
    since = (parse_ts(now) - datetime.timedelta(days=365)).strftime(TS)
    evs = events_for([gid], [num], since, broadcasts([gid], since)).get((gid, num), [])
    try:
        evs = pbt_fix_times(groups[gid], num, evs, now)
    except Exception:
        pass    # PBT unreachable: fall back to TouchPoint's times
    cands = people_on([num], [gid]).get(num, [])
    return groups[gid], evs, cands, thread_facts(gid, evs, cands)


def pbt_fix_times(g, num, evs, now):
    """TouchPoint stamps a text sent outside it (PBT's website, auto-replies)
    when PBT's delivery report arrives, which can be minutes late and puts
    the thread out of order. PBT's own report has the real send time and the
    message type, so the last 7 days are corrected from it. Group sends
    (broadcasts) are left out of the request: they can be thousands of rows."""
    end = parse_ts(now)
    zone, tz = church_offset()
    body = {'groupIds': [g['groupId']], 'pageNumber': 1, 'pageSize': 500,
            'fromDate': (end - datetime.timedelta(days=7)).strftime('%Y-%m-%d'), 'toDate': end.strftime('%Y-%m-%d'),
            'messageTypes': ['TwoWay', 'Twoway', 'Unicast', 'AutoResponse', 'StopResponse', 'StopAllResponse',
                             'HelpResponse', 'OptInResponse', 'KeywordResponse', 'Keyword', 'ReservedKeyword',
                             'SystemMessage', 'Notification', 'Deact']}
    if tz:
        body['timeZoneOffset'] = tz
    pbt_out = []
    for _ in range(4):
        res = pbt_post('/reports/message-reports', body)
        if res.get('isError'):
            if res.get('appCode') == 're.1006':     # PBT still on yesterday's date
                body['toDate'] = (end - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
                continue
            break                                   # includes re.1015, "no messages"
        data = res.get('data') or {}
        for r in data.get('data') or []:
            if digits10(r.get('subscriberNo')) == num and (r.get('direction') or '').lower().startswith('out'):
                t = pbt_ts(r.get('timestamp'))
                if t:
                    pbt_out.append({'t': t, 'body': r.get('message') or '',
                                    'type': (r.get('messageType') or '').lower(), 'used': False})
        if body['pageNumber'] >= (data.get('totalPages') or 1):
            break
        body['pageNumber'] += 1
    if not pbt_out:
        return evs
    for e in evs:
        if e['dir'] != 'out' or e.get('timeFixed'):
            continue
        best, best_m = None, None
        for p in pbt_out:
            if p['used'] or norm(p['body']) != norm(e['body']):
                continue
            m = minutes_between(p['t'], e['t'])    # TouchPoint time minus PBT time
            if m is not None and -2 <= m <= 60 and (best_m is None or abs(m) < abs(best_m)):
                best, best_m = p, m
        if best:
            best['used'] = True
            e['t'], e['timeFixed'] = best['t'], True
            if best['type'] in AUTO_TYPES:
                e['auto'] = True
            elif best['type'] in GROUP_TYPES:
                e['broadcast'] = True
    evs.sort(key=lambda x: x['t'])
    return evs


def pbt_group_texts(g, num, now, evs):
    """Group texts this number got in the last 7 days that TouchPoint never
    saw: ones sent from PoweredByText's own app are not reported back to
    TouchPoint, so a reply like "Yes" would otherwise have nothing above it.
    The report has one row per recipient, so it is read only when a
    conversation is opened (the page keeps the result), and only the texts
    this number actually received are returned. Nothing is stored."""
    end = parse_ts(now)
    zone, tz = church_offset()
    body = {'groupIds': [g['groupId']], 'pageNumber': 1, 'pageSize': 500,
            'fromDate': (end - datetime.timedelta(days=7)).strftime('%Y-%m-%d'), 'toDate': end.strftime('%Y-%m-%d'),
            'messageTypes': ['Broadcast', 'BroadcastMessage', 'ScheduledBroadcast']}
    if tz:
        body['timeZoneOffset'] = tz
    found = []
    for _ in range(14):                 # up to 7,000 recipient rows
        res = pbt_post('/reports/message-reports', body)
        if res.get('isError'):
            if res.get('appCode') == 're.1006' and body['toDate'] == end.strftime('%Y-%m-%d'):
                body['toDate'] = (end - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
                continue
            break                       # includes re.1015, "no messages"
        data = res.get('data') or {}
        for r in data.get('data') or []:
            if digits10(r.get('subscriberNo')) != num or not (r.get('direction') or '').lower().startswith('out'):
                continue
            t = pbt_ts(r.get('timestamp'))
            if t:
                found.append({'t': t, 'body': r.get('message') or ''})
        if body['pageNumber'] >= (data.get('totalPages') or 1):
            break
        body['pageNumber'] += 1
    out = []
    for f in sorted(found, key=lambda x: x['t']):
        # Skip any TouchPoint already shows (sent from this inbox or TouchPoint).
        dup = False
        for e in evs + out:
            m = minutes_between(e['t'], f['t'])
            if e['dir'] == 'out' and norm(e['body']) == norm(f['body']) and m is not None and abs(m) <= 60:
                dup = True
                break
        if not dup:
            out.append({'dir': 'out', 't': f['t'], 'body': f['body']})
    return [{'t': o['t'], 'body': o['body']} for o in out]


def quiet_check(facts, now):
    hour = parse_ts(now).hour
    if QUIET_END <= hour < QUIET_START:
        return ''
    li = facts['lastIn']
    m = minutes_between(li['t'], now) if li else None
    if m is not None and m <= QUIET_GRACE_MIN:
        return ''
    return 'It is quiet hours (9 p.m. to 8 a.m.). Texts can go out now only to someone who texted in within the last 2 hours.'


def act_thread():
    gid = int(_param('gid', '0'))
    num = digits10(_param('num'))
    now = client_now()
    g, evs, cands, f = thread_context(gid, num, now)
    tid = thread_ids([(gid, num)]).get((gid, num), '') if evs else ''
    if tid and ME and _param('claim') in ('1', 'read'):
        # Opening it (or staying on it while new texts arrive) marks it read for me.
        reads = load_reads()
        reads[tid] = now
        save_reads(reads, now)
    st = load_state().get(tid, {}) if tid else {}
    group_texts = None
    if _param('gt') == '1':
        try:
            group_texts = pbt_group_texts(g, num, now, evs)
        except Exception:
            group_texts = []    # PBT unreachable: show what TouchPoint has
    reason = send_block_reason(g)
    can = not reason
    if reason:
        pass
    elif f['opted']:
        can, reason = False, 'This person texted STOP, so they cannot be texted from this group.'
    elif f['blocked']:
        can, reason = False, f['blocked']
    elif not evs and not (f['target'] and f['target']['subs'].get(gid, (False, False))[0]):
        can, reason = False, 'This person is not subscribed to this group, so a new text cannot start here.'
    # Other groups of mine this person is subscribed to, for the From line.
    also = []
    if f['target']:
        mine = my_groups()
        for og, (sub, opt) in f['target']['subs'].items():
            if og != gid and sub and not opt and og in mine and can_send_in(mine[og]):
                also.append({'gid': og, 'name': mine[og]['name']})
    return {'success': True, 'group': g['name'], 'gid': gid, 'num': num, 'alsoIn': sorted(also, key=lambda x: x['name']),
            'events': [{'dir': e['dir'], 't': e['t'], 'body': e['body'],
                        'staff': e.get('staff') or ('PoweredByText' if e.get('pbt') else ''),
                        'broadcast': bool(e.get('broadcast')), 'status': e.get('status', ''),
                        'reaction': bool(e.get('reaction')), 'stop': bool(e.get('stop')),
                        'auto': bool(e.get('auto'))} for e in evs[-200:]],
            'people': [{'pid': c['pid'], 'name': c['name'], 'age': c['age'], 'adult': c['adult'],
                        'receive': c['receive'], 'opted': c['subs'].get(gid, (False, False))[1],
                        'target': bool(f['target'] and f['target']['pid'] == c['pid'])} for c in cands],
            'who': f['who']['name'] if f['who'] else '',
            'canSend': can, 'reason': reason, 'quiet': quiet_check(f, now),
            'via': 'touchpoint' if f['target'] else 'pbt',
            'done': st.get('done') or '', 'doneBy': st.get('doneBy') or '', 'tid': tid,
            'groupTexts': group_texts}


def act_send():
    gid = int(_param('gid', '0'))
    num = digits10(_param('num'))
    msg = _param('msg')
    now = client_now()
    if not msg:
        raise Exception('The message is empty.')
    if len(msg) > 1600:
        raise Exception('Texts can be at most 1,600 characters.')
    g, evs, cands, f = thread_context(gid, num, now)
    if send_block_reason(g):
        raise Exception(send_block_reason(g))
    if f['opted']:
        raise Exception('This person texted STOP, so they cannot be texted from this group.')
    if f['blocked']:
        raise Exception(f['blocked'])
    if not evs and not (f['target'] and f['target']['subs'].get(gid, (False, False))[0]):
        raise Exception('This person is not subscribed to this group, so a new text cannot start here.')
    quiet = quiet_check(f, now)
    if quiet:
        raise Exception(quiet)
    check_emoji(msg, False)
    media = parse_media(_param('media'))
    msg, note = shorten_links(g, msg)
    me = my_name()
    if f['target'] and not media:
        # TouchPoint's own sender: logs it under this staff member and refuses
        # anyone with texts turned off or opted out of the group.
        model.SendSms(str(f['target']['pid']), gid, SEND_TITLE, msg)
        via = 'TouchPoint'
    elif cands and not f['target']:
        raise Exception('Nobody on this number can be texted.')
    else:
        # A picture message has to go through PBT directly (TouchPoint's
        # sender is text only), so TouchPoint's texts-off rule is checked here.
        if f['target'] and not f['target'].get('receive', True):
            raise Exception('This person has texts turned off in TouchPoint.')
        fb = media_fallback(g, _param('fb') or msg, media, _param('fblink') != '0') if media else None
        pbt_send(g, num, msg, media, fb)
        log_sent({'t': now, 'gid': gid, 'b': fingerprint(msg), 'pid': ME, 'by': (me.Name if me else ''),
                  'mms': bool(media)})
        via = 'PoweredByText'
    return {'success': True, 'via': via, 'note': note}


def act_done():
    gid = int(_param('gid', '0'))
    num = digits10(_param('num'))
    now = client_now()
    if gid not in my_groups():
        raise Exception("You don't have rights to that texting group.")
    tid = thread_ids([(gid, num)]).get((gid, num))
    if not tid:
        raise Exception('There is nothing in this conversation to mark yet.')
    st = load_state()
    s = st.setdefault(tid, {})
    if _param('undo') == '1':
        s.pop('done', None)
        s.pop('doneBy', None)
    else:
        me = my_name()
        s['done'], s['doneBy'] = now, (me.First if me else '')
    save_state(st, now)
    return {'success': True}


def act_mark_read():
    """Mark many conversations read for me (e.g. a day of RSVP answers)."""
    now = client_now()
    groups = my_groups()
    pairs = []
    for k in _param('keys').split(','):
        m = re.match(r'^(\d+)\|(\d{10})$', k.strip())
        if m and int(m.group(1)) in groups:
            pairs.append((int(m.group(1)), m.group(2)))
    tids = thread_ids(pairs)
    reads = load_reads()
    for t in tids.values():
        reads[t] = now
    save_reads(reads, now)
    return {'success': True, 'marked': len(tids)}


def act_find():
    """People a new text can go to: in TouchPoint, with a cell, texts on, and
    subscribed (not opted out) to at least one of my groups."""
    term = _param('term')
    groups = dict((k, v) for k, v in my_groups().items() if can_send_in(v))
    if len(term) < 2 or not groups:
        return {'success': True, 'people': []}
    digits = re.sub(r'\D', '', term)
    if len(digits) >= 4:
        cond = "p.CellPhone LIKE '%%%s%%'" % digits
    else:
        # Every word must appear somewhere in the first, goes-by, nick or last
        # name, in any order: "sar whit", "whit sar" and "aby" all work.
        words = [w for w in re.split(r'\s+', re.sub(r"[^\w' -]", ' ', term, flags=re.U)) if w][:4]
        if not words:
            return {'success': True, 'people': []}
        cond = ' AND '.join(
            "(p.FirstName LIKE N'%%{0}%%' OR ISNULL(p.PreferredName, '') LIKE N'%%{0}%%' OR ISNULL(p.NickName, '') LIKE N'%%{0}%%' "
            "OR p.LastName LIKE N'%%{0}%%' OR ISNULL(p.MaidenName, '') LIKE N'%%{0}%%')".format(w.replace("'", "''")) for w in words)
    rows = q.QuerySql("""
        SELECT TOP 25 p.PeopleId, ISNULL(NULLIF(p.PreferredName, ''), p.FirstName) + ' ' + p.LastName N, p.FirstName Legal, p.Age, p.CellPhone,
               ISNULL(p.ReceiveSMS, 0) R
        FROM People p WITH (NOLOCK)
        WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0 AND LEN(p.CellPhone) = 10 AND %s
        ORDER BY p.Name2""" % cond)
    pids = [r.PeopleId for r in rows]
    subs = {}
    if pids:
        for s in q.QuerySql("""
                SELECT ToPeopleId, FromGroup FROM SMSGroupSubscriptions WITH (NOLOCK)
                WHERE ToPeopleId IN (%s) AND FromGroup IN (%s)
                  AND ISNULL(IsSubscribed, 0) = 1 AND ISNULL(IsOptedOut, 0) = 0""" % (sql_ints(pids), sql_ints(groups.keys()))):
            subs.setdefault(s.ToPeopleId, []).append({'gid': s.FromGroup, 'name': groups[s.FromGroup]['name']})
        # People who joined by texting a keyword are subscribed in PBT but not in
        # TouchPoint's table. Having texted the group counts, unless the last
        # STOP/START word they sent it was STOP.
        cells = dict((r.CellPhone, r.PeopleId) for r in rows)
        state = {}
        for r in q.QuerySql("""
                SELECT r.ToGroupId G, %s Num, r.Body, r.Action
                FROM SmsReceived r WITH (NOLOCK)
                WHERE r.ToGroupId IN (%s) AND %s IN (%s)
                ORDER BY r.DateReceived""" % (NUM_SQL, sql_ints(groups.keys()), NUM_SQL, sql_strs(cells.keys()))):
            k = (r.Num, r.G)
            if is_stop(r.Body) or (r.Action or '').upper() in ('STOP', 'STOPALL'):
                state[k] = False
            elif k not in state or is_start(r.Body):
                state[k] = True
        for (num, gid), ok in state.items():
            pid = cells.get(num)
            if ok and pid and not any(x['gid'] == gid for x in subs.get(pid, [])):
                subs.setdefault(pid, []).append({'gid': gid, 'name': groups[gid]['name']})
    return {'success': True, 'people': [
        {'pid': r.PeopleId, 'name': r.N, 'legal': r.Legal if r.N and r.Legal and not r.N.startswith(r.Legal + ' ') else '',
         'age': r.Age, 'num': r.CellPhone, 'receive': bool(r.R),
         'groups': sorted(subs.get(r.PeopleId, []), key=lambda x: x['name'])}
        # Only people this staff member could actually text are listed.
        for r in rows if r.R and subs.get(r.PeopleId)]}


# ---------------------------------------------------------------- PBT report test (read only)
# Answers "could TPxi Go run this inbox straight from PBT?" PBT's message
# report is read for every PBT group and compared with what TouchPoint
# stored from its webhook: does the report carry incoming texts with the
# number and body, which fields are filled in, and how far apart the two
# clocks are (time zone plus any delay). Nothing is sent or changed.
REPORT_FIELDS = ['timestamp', 'direction', 'messageType', 'subscriberNo', 'message', 'groupId', 'groupName',
                 'subscriberFirstName', 'subscriberLastName', 'statusCodeDescription', 'mediaUrl',
                 'referenceId', 'utilizedAmount', 'customParams']


def pbt_post(path, obj):
    tok = pbt_token()
    raw = model.RestPost(_conn['api'] + path, {"Authorization": "Bearer " + tok, "Accept": "application/json",
                                              "Content-Type": "application/json"}, safe_json(obj))
    try:
        return json.loads(raw) if raw and raw.strip() else {}
    except Exception:
        raise Exception("PoweredByText reply was not understood: %s" % (raw or '(empty)')[:300])


def church_offset():
    """TouchPoint's LocalTimeZone setting (a Windows zone name such as
    'Central Standard Time') turned into PBT's '-05:00' form for right now,
    daylight saving included."""
    name = model.Setting("LocalTimeZone", "Central Standard Time") or "Central Standard Time"
    try:
        import System
        span = System.TimeZoneInfo.FindSystemTimeZoneById(name).GetUtcOffset(System.DateTime.UtcNow)
        m = int(span.TotalMinutes)
        return name, '%s%02d:%02d' % ('-' if m < 0 else '+', abs(m) // 60, abs(m) % 60)
    except Exception:
        return name, None


def pbt_ts(v):
    """'2026-10-24T14:05:31.123Z' -> '2026-10-24 14:05:31' (zone dropped; the
    clock comparison below works out the zone)."""
    v = (v or '').replace('T', ' ')
    return v[:19] if parse_ts(v[:19]) else ''


def act_pbt_test():
    if not model.UserIsInRole('Admin'):
        raise Exception('Admin only.')
    now = client_now()
    hours = int(_param('hours', '24')) if _param('hours', '24').isdigit() else 24
    hours = max(1, min(hours, 24 * 60))
    end = parse_ts(now)
    start = end - datetime.timedelta(hours=hours)
    groups = {}
    for r in q.QuerySql("""
            SELECT ID, Name, ExternalId FROM SMSGroups
            WHERE GroupType = 'PoweredByText' AND IsDeleted = 0 AND ISNULL(isDisabled, 0) = 0
              AND ISNULL(OneTimeCodeFlag, 0) = 0 AND ExternalId IS NOT NULL"""):
        groups[str(r.ExternalId)] = {'id': r.ID, 'name': r.Name}
    body = {'groupIds': list(groups.keys()), 'pageNumber': 1, 'pageSize': 200,
            'fromDate': start.strftime('%Y-%m-%d'), 'toDate': end.strftime('%Y-%m-%d')}
    zone, church_tz = church_offset()
    tz = _param('tz') or church_tz
    if tz:
        body['timeZoneOffset'] = tz
    recs, pages, total, calls = [], 0, None, []
    while pages < 15:
        t0 = datetime.datetime.now()
        res = pbt_post('/reports/message-reports', body)
        ms = int((datetime.datetime.now() - t0).total_seconds() * 1000)
        calls.append(ms)
        if res.get('isError') and res.get('appCode') == 're.1006' and body['toDate'] == end.strftime('%Y-%m-%d'):
            # "ToDate should not be greater than current date": PBT's clock
            # (likely UTC or its own zone) is still on the day before ours.
            body['toDate'] = (end - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            if body['fromDate'] > body['toDate']:
                body['fromDate'] = body['toDate']
            continue
        if res.get('isError'):
            raise Exception('PoweredByText: %s %s' % (res.get('appCode') or '', res.get('message') or ''))
        data = res.get('data') or {}
        rows = data.get('data') or []
        total = data.get('totalCount', total)
        recs.extend(rows)
        pages += 1
        if not rows or body['pageNumber'] >= (data.get('totalPages') or 1):
            break
        body['pageNumber'] += 1

    # Field coverage and counts.
    cover = {}
    for f in REPORT_FIELDS:
        cover[f] = sum(1 for r in recs if r.get(f) not in (None, '', [], {}))
    by_dir, by_type = {}, {}
    for r in recs:
        d = r.get('direction') or '(blank)'
        by_dir[d] = by_dir.get(d, 0) + 1
        k = '%s / %s' % (d, r.get('messageType') or '(blank)')
        by_type[k] = by_type.get(k, 0) + 1

    # Compare incoming texts with what TouchPoint stored.
    inb = [r for r in recs if (r.get('direction') or '').lower().startswith('in')]
    tp = {}
    for r in q.QuerySql("""
            SELECT CONVERT(varchar(19), r.DateReceived, 120) T, r.ToGroupId G, r.Body, %s Num
            FROM SmsReceived r WITH (NOLOCK)
            WHERE r.DateReceived >= '%s'""" % (NUM_SQL, (start - datetime.timedelta(days=1)).strftime(TS))):
        tp.setdefault((r.G, r.Num, norm(r.Body)), []).append(r.T)
    diffs, matched, missing_tp, samples_missing = [], 0, 0, []
    for r in inb:
        g = groups.get(str(r.get('groupId')))
        t = pbt_ts(r.get('timestamp'))
        key = (g['id'] if g else None, digits10(r.get('subscriberNo')), norm(r.get('message')))
        cands = tp.get(key) or []
        if cands and t:
            best = min(cands, key=lambda x: abs(minutes_between(x, t) or 0))
            diffs.append(round(minutes_between(best, t), 1))
            matched += 1
        else:
            missing_tp += 1
            if len(samples_missing) < 5:
                samples_missing.append({'t': r.get('timestamp'), 'group': (g or {}).get('name') or r.get('groupName'),
                                        'num': digits10(r.get('subscriberNo')), 'msg': (r.get('message') or '')[:80]})
    diffs.sort()
    tp_window = sum(1 for k, v in tp.items() for t in v if t >= start.strftime(TS))
    sample = []
    for want in ('in', 'out'):
        for r in recs:
            if (r.get('direction') or '').lower().startswith(want):
                sample.append(dict((f, r.get(f)) for f in REPORT_FIELDS))
                if len([x for x in sample if (x.get('direction') or '').lower().startswith(want)]) >= 3:
                    break
    return {'success': True, 'zone': zone, 'zoneOffset': church_tz, 'api': _conn.get('api'), 'request': dict((k, v) for k, v in body.items() if k != 'groupIds'),
            'groupCount': len(groups), 'records': len(recs), 'totalCount': total, 'pages': pages,
            'callMs': calls, 'byDirection': by_dir, 'byType': by_type, 'coverage': cover,
            'inbound': len(inb), 'matched': matched, 'notInTouchPoint': missing_tp,
            'tpInbound': tp_window, 'missingSamples': samples_missing,
            'clockDiffMin': {'n': len(diffs), 'min': diffs[0] if diffs else None,
                             'median': diffs[len(diffs) // 2] if diffs else None, 'max': diffs[-1] if diffs else None},
            'sample': sample}


def act_pbt_speed():
    """Times the same kind of read several ways to find where the seconds go:
    sign-in, a tiny ordinary call, then the report narrowed step by step."""
    if not model.UserIsInRole('Admin'):
        raise Exception('Admin only.')
    now = parse_ts(client_now())
    today = now.strftime('%Y-%m-%d')
    yesterday = (now - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    zone, tz = church_offset()
    gids, one = [], None
    for r in q.QuerySql("""
            SELECT ExternalId, ExternalId2, Name FROM SMSGroups
            WHERE GroupType = 'PoweredByText' AND IsDeleted = 0 AND ISNULL(isDisabled, 0) = 0
              AND ISNULL(OneTimeCodeFlag, 0) = 0 AND ExternalId IS NOT NULL ORDER BY Name"""):
        gids.append(str(r.ExternalId))
        if one is None or r.Name == 'FBCHville':
            one = (str(r.ExternalId), str(r.ExternalId2), r.Name)
    out = []

    def timed(label, fn):
        t0 = datetime.datetime.now()
        try:
            n = fn()
            err = ''
        except Exception as e:
            n, err = None, unicode(e)[:200]
        out.append({'label': label, 'ms': int((datetime.datetime.now() - t0).total_seconds() * 1000),
                    'records': n, 'error': err})

    def report(groups, frm, types=None, size=200):
        body = {'groupIds': groups, 'pageNumber': 1, 'pageSize': size, 'fromDate': frm, 'toDate': today}
        if tz:
            body['timeZoneOffset'] = tz
        if types:
            body['messageTypes'] = types
        res = pbt_post('/reports/message-reports', body)
        if res.get('isError'):
            raise Exception('%s %s' % (res.get('appCode') or '', res.get('message') or ''))
        return (res.get('data') or {}).get('totalCount')

    def first_call():
        _conn.pop('token', None)
        pbt_token()
        return None
    timed('Sign in (token)', first_call)
    timed('Small ordinary call: 1 subscriber of %s' % one[2], lambda: len(
        (pbt_get_raw('/orgunit/%s/groups/%s/subscribers?pageNumber=1&pageSize=1' % (one[1], one[0])).get('data') or {}).get('data') or []))
    timed('Report: all %d groups, yesterday and today' % len(gids), lambda: report(gids, yesterday))
    timed('Report: all %d groups, today only' % len(gids), lambda: report(gids, today))
    timed('Report: all groups, today, incoming types only', lambda: report(gids, today, ['TwoWay', 'Keyword', 'ReservedKeyword']))
    timed('Report: 1 group (%s), today only' % one[2], lambda: report([one[0]], today))
    timed('Report: 1 group, today, page size 10', lambda: report([one[0]], today, None, 10))
    timed('Report: all groups, today, again (warm)', lambda: report(gids, today))
    return {'success': True, 'zone': zone, 'zoneOffset': tz, 'api': _conn.get('api'), 'steps': out}


def pbt_get_raw(path):
    tok = pbt_token()
    raw = model.RestGet(_conn['api'] + path, {"Authorization": "Bearer " + tok, "Accept": "application/json"})
    try:
        return json.loads(raw)
    except Exception:
        raise Exception("Not JSON: %s" % (raw or '(empty)')[:200])


PBT_TEST_PAGE = """
<div id="pt" style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:1100px">
  <p><a href="?">&larr; Back to the inbox</a></p>
  <p>Reads PoweredByText's message report for every texting group and compares it with what TouchPoint stored.
  It only reads: nothing is sent or changed. Admin only.</p>
  <p>Look back <select id="hours"><option value="1">1 hour</option><option value="6">6 hours</option>
  <option value="24" selected>24 hours</option><option value="168">7 days</option></select>
  Time zone offset <input id="tz" placeholder="from TouchPoint" style="width:110px"> <span style="color:#6b7280;font-size:12px">(blank uses TouchPoint's LocalTimeZone; PBT wants the -05:00 form)</span>
  <button id="go" style="padding:6px 12px">Run test</button> <button id="speed" style="padding:6px 12px">Speed check</button> <span id="st"></span></p>
  <div id="out"></div>
</div>
<script>
(function(){
  var base = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
  function pad(n){ return (n < 10 ? '0' : '') + n; }
  function nowStr(){ var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate()) + ' ' +
    pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds()); }
  function esc(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
  document.getElementById('speed').onclick = function(){
    var st = document.getElementById('st'), out = document.getElementById('out'), btn = this;
    btn.disabled = true; st.textContent = 'Timing PoweredByText (about 30 seconds)...'; out.innerHTML = '';
    var x = new XMLHttpRequest();
    x.open('POST', base, true);
    x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    x.onload = function(){
      btn.disabled = false; var r;
      try { r = JSON.parse(x.responseText); } catch(e) { st.textContent = 'Unexpected reply: ' + x.responseText.substr(0, 200); return; }
      if (!r.success) { st.textContent = r.message; return; }
      st.textContent = 'Done. Time zone: ' + r.zone + ' (' + (r.zoneOffset || 'unknown') + ')';
      out.innerHTML = '<h3>Speed check</h3><table style="border-collapse:collapse;font-size:13px">' +
        '<tr><th style="text-align:left;padding:2px 12px 2px 0">Step</th><th style="text-align:right;padding:2px 12px">ms</th><th style="text-align:right;padding:2px 12px">Records</th><th></th></tr>' +
        r.steps.map(function(s){ return '<tr><td style="padding:2px 12px 2px 0">' + esc(s.label) + '</td><td style="text-align:right;padding:2px 12px">' + s.ms +
          '</td><td style="text-align:right;padding:2px 12px">' + (s.records == null ? '' : s.records) + '</td><td style="color:#a12a22">' + esc(s.error) + '</td></tr>'; }).join('') + '</table>';
    };
    x.onerror = function(){ btn.disabled = false; st.textContent = 'Network error'; };
    x.send('action=pbt_speed&now=' + encodeURIComponent(nowStr()));
  };
  function table(obj){ return '<table style="border-collapse:collapse;font-size:13px">' + Object.keys(obj).map(function(k){
    return '<tr><td style="padding:2px 12px 2px 0;color:#555">' + esc(k) + '</td><td>' + esc(obj[k]) + '</td></tr>'; }).join('') + '</table>'; }
  document.getElementById('go').onclick = function(){
    var st = document.getElementById('st'), out = document.getElementById('out'), btn = this;
    btn.disabled = true; st.textContent = 'Reading PoweredByText...'; out.innerHTML = '';
    var data = {action:'pbt_test', hours:document.getElementById('hours').value, tz:document.getElementById('tz').value, now:nowStr()};
    var x = new XMLHttpRequest();
    x.open('POST', base, true);
    x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    x.onload = function(){
      btn.disabled = false; var r;
      try { r = JSON.parse(x.responseText); } catch(e) { st.textContent = 'Unexpected reply: ' + x.responseText.substr(0, 200); return; }
      if (!r.success) { st.textContent = r.message; return; }
      st.textContent = 'Done.';
      var c = r.clockDiffMin, h = [];
      h.push('<h3>What came back</h3>' + table({'TouchPoint time zone': r.zone + ' (' + (r.zoneOffset || 'unknown') + ')', 'API': r.api, 'Groups asked for': r.groupCount,
        'Records read': r.records + ' of ' + r.totalCount + ' (' + r.pages + ' page(s))',
        'Time per request (ms)': r.callMs.join(', '), 'Request': JSON.stringify(r.request)}));
      h.push('<h3>By direction</h3>' + table(r.byDirection) + '<h3>By direction and type</h3>' + table(r.byType));
      var cov = {}; Object.keys(r.coverage).forEach(function(k){ cov[k] = r.coverage[k] + ' of ' + r.records; });
      h.push('<h3>Fields filled in</h3>' + table(cov));
      h.push('<h3>Incoming texts compared with TouchPoint</h3>' + table({
        'Incoming in the PBT report': r.inbound,
        'Matched to a TouchPoint record (same group, number, text)': r.matched,
        'In PBT but not found in TouchPoint': r.notInTouchPoint,
        'Incoming stored in TouchPoint in this window': r.tpInbound,
        'PBT time minus TouchPoint time, minutes (min / median / max)': c.n ? c.min + ' / ' + c.median + ' / ' + c.max : 'no matches'}));
      if (r.missingSamples.length) h.push('<h3>Examples not found in TouchPoint</h3><pre style="font-size:12px">' + esc(JSON.stringify(r.missingSamples, null, 2)) + '</pre>');
      h.push('<h3>Sample records</h3><pre style="font-size:12px;white-space:pre-wrap">' + esc(JSON.stringify(r.sample, null, 2)) + '</pre>');
      out.innerHTML = h.join('');
    };
    x.onerror = function(){ btn.disabled = false; st.textContent = 'Network error'; };
    x.send(Object.keys(data).map(function(k){ return k + '=' + encodeURIComponent(data[k]); }).join('&'));
  };
})();
</script>"""



# ================================================================ admin tools (was TPxi_PBTNameSync)
# PoweredByText (PBT) shows a texting subscriber by name only when its
# contact record has one. TouchPoint sends a name to PBT just once, at the
# moment TouchPoint itself opts a person in. Everyone who joined by keyword,
# was added in PBT, came over from MinistryByText, or changed their name
# later shows up in PBT as a bare number.
#
# This page reads every subscriber in the PBT groups you pick, finds who
# owns each number in TouchPoint, shows the name it would set and why, and
# sets it for the rows you choose. It can also run every morning.
#
# WHO OWNS A NUMBER (people whose Cell Phone is that number, not deceased,
# not archived):
#   1. Adults beat children. A parent's number copied onto a child's
#      record goes to the parent.
#   2. One adult left: that adult.
#   3. Several adults in ONE family: the head of household.
#   4. Adults in DIFFERENT families: no guess, listed for cleanup.
#   5. Only children: a teen (TEEN_AGE+) is taken as the owner if there is
#      exactly one; a younger child is listed for cleanup, because the
#      number almost certainly belongs to a parent whose record lacks it.
#
# WHAT A SYNC DOES with the name already in PBT:
#   Will set      PBT has no name.                         -> written
#   Will fix      PBT has the name of someone ELSE on this
#                 number (usually the child).              -> written
#   Already right PBT already has the chosen name.         -> nothing
#   Keep          PBT has a name that matches nobody on
#                 this number, so a person typed it in PBT. -> never written
#   Review        No safe owner (rules 4 and 5).           -> nothing
#   Not found     No TouchPoint person has this cell.      -> nothing
#
# The decision is made again on the server at the moment of writing, from a
# fresh read of PBT, so a row that turned into Keep since the preview is
# skipped. Names go to PBT's contact record for the brand (POST /contacts),
# never through the group subscribe call, which would send an opt-in text.
# The TouchPoint PeopleId is stored on the contact as the custom field
# TouchPointPeopleId.
#
# EVERY CHANGE IS LOGGED to Special Content text TPxi_PBTNameSyncLog (newest
# LOG_KEEP entries) with the name before and after, so it can be undone.
#
# NAME CHANGES ARE ON HOLD (NAME_WRITES below): PBT has no way yet for an
# outside system to set a contact's name. The Names tab still shows who
# owns each number and which TouchPoint records need fixing.
#
# MORNING BATCH: turning on "Update names every morning" on the page adds a
# small managed block to the MorningBatch script (and turning it off removes
# it). A morning run that would change more than AUTO_MAX names changes
# nothing and logs why, so a data problem cannot rename the whole list.
#
# DELIVERY AND USAGE TABS (read only): PBT's message report has one row per
# recipient per text, with the time it went out and its delivery status.
# Delivery groups those rows into sends, shows when each really went out and
# how it went, and matches texts that came from TouchPoint (by the reference
# id TouchPoint stores in SMSItems.PbtReferenceId, else by text and time) to
# the time TouchPoint was asked to send them. PBT does not report the
# scheduled time of a text scheduled inside PBT, only when it went out.
# Usage counts the same rows by month and group. Segments are ESTIMATED
# from message length (160/153 characters, 70/67 when the text has
# characters outside the basic texting alphabet, such as emoji).
#
# CREDENTIALS: reuses the PBT client id and secret TouchPoint already
# stores in Settings (PBTClientID, PBTClientSecret). Nothing to configure.
#

PAGE_SIZE = 1000          # PBT maximum
TEEN_AGE = 13             # an only-child number at this age or older is the teen's own
AUTO_MAX = 300            # a morning run that would change more than this changes nothing
# PBT does not yet let an outside system set a contact's name, so changing
# names is off: no Change buttons, no morning renames, and the apply action
# refuses. The Names tab is a report of what is off until PBT adds it.
NAME_WRITES = False
LOG_KEEP = 3000           # log entries kept
PEOPLEID_PARAM = 'TouchPointPeopleId'

SETTINGS_CONTENT = 'TPxi_PBTNameSyncSettings'
USAGE_CONTENT = 'TPxi_PBTNameSyncUsage'     # finished months, counted once (counts only, no numbers)
LOG_CONTENT = 'TPxi_PBTNameSyncLog'
MB_START = "# >>> TPxi_PBTNameSync start (managed by app, do not edit) >>>"
MB_END = "# <<< TPxi_PBTNameSync end <<<"

# Groups that are TouchPoint plumbing, not people lists.
# ---------------------------------------------------------------- JSON
# IronPython json.dumps breaks on non-ASCII names; see CLAUDE.md.
# ---------------------------------------------------------------- helpers
def now_str():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M')


def pbt_groups(include_system=False):
    """PBT groups known to TouchPoint. include_system adds the one-time code
    and system groups: no one needs a name there, but their texts still use
    credits, so the Usage and Delivery tabs count them."""
    rows = q.QuerySql("""
        SELECT ID, Name, ExternalId, ExternalId2
        FROM SMSGroups
        WHERE GroupType = 'PoweredByText' AND IsDeleted = 0
          AND ISNULL(isDisabled, 0) = 0
          %s
          AND ExternalId IS NOT NULL AND ExternalId2 IS NOT NULL
        ORDER BY Name""" % ('' if include_system else 'AND ISNULL(OneTimeCodeFlag, 0) = 0'))
    out = []
    for r in rows:
        if SKIP_NAME.search(r.Name or '') and not include_system:
            continue
        out.append({'id': r.ID, 'name': r.Name, 'groupId': r.ExternalId, 'orgunitId': r.ExternalId2})
    return out


# ---------------------------------------------------------------- PBT API


def pbt_get(path):
    tok = pbt_token()
    raw = model.RestGet(_conn['api'] + path, {"Authorization": "Bearer " + tok,
                                             "Accept": "application/json"})
    try:
        return json.loads(raw)
    except Exception:
        raise Exception("PBT returned something that is not JSON for %s: %s" % (path, (raw or '(empty)')[:200]))


def pbt_subscribers(orgunit_id, group_id):
    subs = []
    page = 1
    # The current API calls a brand an "orgunit"; TouchPoint's older client says "campuses".
    kinds = ['orgunit', 'campuses']
    while True:
        try:
            res = pbt_get("/%s/%s/groups/%s/subscribers?pageNumber=%d&pageSize=%d"
                          % (kinds[0], orgunit_id, group_id, page, PAGE_SIZE))
            if res.get('isError') or not isinstance(res.get('data'), dict):
                raise Exception("PBT error: %s %s" % (res.get('appCode') or '', res.get('message') or ''))
        except Exception:
            if page == 1 and len(kinds) > 1:
                kinds.pop(0)
                continue
            raise
        data = res.get('data') or {}
        rows = data.get('data') or []
        subs.extend(rows)
        total_pages = data.get('totalPages') or 1
        if page >= total_pages or not rows:
            break
        page += 1
    return subs


def pbt_set_name(orgunit_id, mobile, first, last, pid):
    """Writes the name on the brand's contact record. Returns a note ('' when all went in)."""
    body = {'firstName': first, 'lastName': last, 'mobileNo': mobile}
    if pid and not _conn.get('no_custom'):
        body['customParams'] = [{'name': PEOPLEID_PARAM, 'value': str(pid)}]
    res = pbt_post("/orgunit/%s/contacts" % orgunit_id, body)
    if res.get('isError') and 'customParams' in body:
        # A brand may not allow a custom field we have not set up there. The
        # name matters more than the id, so try once more without it.
        first_msg = res.get('message') or res.get('appCode') or ''
        del body['customParams']
        res = pbt_post("/orgunit/%s/contacts" % orgunit_id, body)
        if not res.get('isError'):
            _conn['no_custom'] = True
            return "PeopleId not stored (%s)" % first_msg
    if res.get('isError'):
        raise Exception(res.get('message') or res.get('appCode') or 'PBT refused the change')
    return ''


# ---------------------------------------------------------------- TouchPoint side
def tp_people_by_cell(numbers):
    """{10-digit number: [person dict, ...]} for active people with that cell phone."""
    found = {}
    nums = sorted(set(n for n in numbers if re.match(r'^\d{10}$', n)))
    for i in range(0, len(nums), 500):
        chunk = nums[i:i + 500]
        sql = """
            SELECT p.PeopleId, p.FamilyId, p.CellPhone, p.FirstName, p.PreferredName,
                   p.LastName, p.Age, p.PositionInFamilyId,
                   f.HeadOfHouseholdId
            FROM People p
            JOIN Families f ON f.FamilyId = p.FamilyId
            WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0
              AND p.CellPhone IN (%s)""" % ','.join("'%s'" % n for n in chunk)
        for r in q.QuerySql(sql):
            found.setdefault(r.CellPhone, []).append({
                'pid': r.PeopleId, 'fam': r.FamilyId,
                'first': (r.PreferredName or r.FirstName) or '',
                'legalFirst': r.FirstName or '', 'last': r.LastName or '',
                'age': r.Age, 'pos': r.PositionInFamilyId,
                'head': r.HeadOfHouseholdId,
            })
    return found


def is_adult(c):
    return c['pos'] != 30 and (c['age'] is None or c['age'] >= 18)


def decide(cands):
    """Returns (owner_first, owner_last, owner_pid or None, reason, review_flag)."""
    if not cands:
        return None, None, None, "No one in TouchPoint has this cell phone", False
    adults = [c for c in cands if is_adult(c)]
    kids = [c for c in cands if not is_adult(c)]
    if len(adults) == 1:
        a = adults[0]
        why = "Only person with this number" if len(cands) == 1 else \
            "Only adult with this number (also on %d child record%s)" % (len(kids), '' if len(kids) == 1 else 's')
        return a['first'], a['last'], a['pid'], why, False
    if len(adults) > 1:
        fams = set(a['fam'] for a in adults)
        if len(fams) > 1:
            return None, None, None, "Shared by adults in %d different families" % len(fams), True
        head = [a for a in adults if a['pid'] == a['head']]
        if head:
            return head[0]['first'], head[0]['last'], head[0]['pid'], \
                "Shared by %d adults in one family; head of household" % len(adults), False
        primary = sorted([a for a in adults if a['pos'] == 10], key=lambda a: -(a['age'] or 0))
        if primary:
            p = primary[0]
            return p['first'], p['last'], p['pid'], \
                "Shared by %d adults in one family; oldest primary adult" % len(adults), False
        return None, None, None, "Shared by %d adults in one family, none a primary adult" % len(adults), True
    teens = [k for k in kids if k['age'] is not None and k['age'] >= TEEN_AGE]
    if len(teens) == 1 and len(kids) == 1:
        t = teens[0]
        return t['first'], t['last'], t['pid'], "Only person with this number (age %d)" % t['age'], False
    if len(teens) == 1:
        t = teens[0]
        return t['first'], t['last'], t['pid'], \
            "Only children have it; the one teen (age %d) is taken as the owner" % t['age'], False
    return None, None, None, "Only children have this number; likely a parent's phone missing from the parent's record", True


def evaluate(g, subs):
    """One row per subscriber with the action a sync would take."""
    nums = [digits10(s.get('mobileNo')) for s in subs]
    people = tp_people_by_cell(nums)
    rows = []
    for s, num in zip(subs, nums):
        cands = people.get(num, [])
        first, last, pid, why, review = decide(cands)
        cur_first = (s.get('firstName') or '').strip()
        cur_last = (s.get('lastName') or '').strip()
        cur = norm(cur_first + ' ' + cur_last)
        if not cands:
            action = 'notfound'
        elif review:
            action = 'review'
        elif not cur:
            action = 'set'
        elif cur == norm(first + ' ' + last):
            action = 'same'
        elif any(cur in (norm(c['first'] + ' ' + c['last']), norm(c['legalFirst'] + ' ' + c['last']))
                 for c in cands):
            action = 'fix'
        else:
            action = 'keep'
        rows.append({
            'num': num, 'mobile': s.get('mobileNo') or '', 'gid': g['id'],
            'status': (s.get('status') or ''),
            'pbt': (cur_first + ' ' + cur_last).strip(),
            'first': first or '', 'last': last or '', 'pid': pid,
            'why': why, 'action': action,
            'cands': [{'pid': c['pid'], 'name': (c['first'] + ' ' + c['last']).strip(), 'age': c['age'],
                       'adult': is_adult(c)} for c in cands],
        })
    return rows


# ---------------------------------------------------------------- writing + log
def write_log(entries):
    if not entries:
        return
    log = load_json_content(LOG_CONTENT, [])
    if not isinstance(log, list):
        log = []
    log = (entries + log)[:LOG_KEEP]
    model.WriteContentText(LOG_CONTENT, safe_json(log), "")


def apply_group(g, only_nums, source, who):
    """Re-reads the group from PBT, re-decides, writes Will set / Will fix rows
    (all of them, or only the numbers in only_nums), then reads PBT again to
    confirm. Returns (results, log entries)."""
    rows = evaluate(g, pbt_subscribers(g['orgunitId'], g['groupId']))
    by_num = {}
    for r in rows:
        by_num.setdefault(r['num'], r)
    targets = only_nums if only_nums is not None else \
        [n for n, r in by_num.items() if r['action'] in ('set', 'fix')]
    results, wrote = [], []
    for num in targets:
        r = by_num.get(num)
        if not r:
            results.append({'num': num, 'result': 'skipped', 'msg': 'No longer in this group'})
            continue
        if r['action'] not in ('set', 'fix'):
            results.append({'num': num, 'result': 'skipped',
                            'msg': 'Now "%s", so it was left alone' % r['action']})
            continue
        try:
            note = pbt_set_name(g['orgunitId'], r['mobile'], r['first'], r['last'], r['pid'])
            wrote.append((r, note))
        except Exception as e:
            results.append({'num': num, 'result': 'error', 'msg': unicode(e)[:200],
                            'before': r['pbt'], 'after': (r['first'] + ' ' + r['last']).strip(), 'pid': r['pid']})
    if wrote:
        after = {}
        try:
            for s in pbt_subscribers(g['orgunitId'], g['groupId']):
                after.setdefault(digits10(s.get('mobileNo')),
                                 ((s.get('firstName') or '') + ' ' + (s.get('lastName') or '')).strip())
        except Exception:
            after = None
        for r, note in wrote:
            new = (r['first'] + ' ' + r['last']).strip()
            ok = after is not None and norm(after.get(r['num'])) == norm(new)
            msg = note if ok else ('PBT accepted it but still shows "%s"' % (after.get(r['num']) or '')
                                   if after is not None else 'Sent; could not read PBT back to confirm')
            if note and not ok:
                msg = note + '. ' + msg
            results.append({'num': r['num'], 'result': 'changed' if ok else 'unconfirmed', 'msg': msg,
                            'before': r['pbt'], 'after': new, 'pid': r['pid']})
    entries = []
    stamp = now_str()
    for x in results:
        if x['result'] == 'skipped':
            continue
        entries.append({'when': stamp, 'by': who, 'source': source, 'group': g['name'],
                        'num': x['num'], 'before': x.get('before', ''), 'after': x.get('after', ''),
                        'pid': x.get('pid'), 'result': x['result'], 'msg': x.get('msg', '')})
    return results, entries


# ---------------------------------------------------------------- settings + morning batch
def load_settings():
    s = load_json_content(SETTINGS_CONTENT, {})
    return s if isinstance(s, dict) else {}


def save_settings(s):
    model.WriteContentText(SETTINGS_CONTENT, safe_json(s), "")


EMOJI_POLICIES = ('allow', 'replies', 'none')
ATTACH_POLICIES = ('off', 'pictures', 'pictures_pdf', 'all')


def send_policy():
    """What staff may put in a text, set by an admin under Settings.
    emoji: allow | replies (allowed, warned in group texts) | none
    attach: off | pictures | pictures_pdf | all (adds small video and audio)"""
    st = load_settings()
    e = st.get('emoji_policy') if st.get('emoji_policy') in EMOJI_POLICIES else 'replies'
    a = st.get('attach_policy') if st.get('attach_policy') in ATTACH_POLICIES else 'pictures_pdf'
    return {'emoji': e, 'attach': a}


def mb_read():
    try:
        return model.PythonContent("MorningBatch") or ""
    except Exception:
        return ""


def mb_remove_block(text):
    if MB_START not in text:
        return text
    head, rest = text.split(MB_START, 1)
    tail = rest.split(MB_END, 1)[1] if MB_END in rest else ''
    return head.rstrip('\n') + '\n' + tail.lstrip('\n')


def mb_install(script_name):
    text = mb_remove_block(mb_read())
    block = '\n'.join([
        MB_START,
        "try:",
        "    Data.run_pbt_name_sync = 'true'",
        "    model.CallScript('%s')" % script_name.replace("'", ""),
        "except Exception as _pbt_e:",
        "    print 'PBT name sync error: ' + str(_pbt_e)",
        MB_END,
    ])
    model.WriteContentPython("MorningBatch", text.rstrip('\n') + '\n\n' + block + '\n')


def mb_uninstall():
    text = mb_read()
    if MB_START in text:
        model.WriteContentPython("MorningBatch", mb_remove_block(text))


def morning_run():
    """Called from MorningBatch. Never raises: one failure here must not stop
    the rest of the morning batch."""
    settings = load_settings()
    if not NAME_WRITES or not settings.get('auto'):
        return
    summary = {'when': now_str(), 'changed': 0, 'unconfirmed': 0, 'errors': 0, 'stopped': ''}
    entries = []
    try:
        chosen = set(settings.get('groups') or [])
        groups = [g for g in pbt_groups() if not chosen or g['id'] in chosen]
        # Plan first: one write per number per brand, and a ceiling on the total.
        plan, seen = [], set()
        for g in groups:
            nums = []
            for r in evaluate(g, pbt_subscribers(g['orgunitId'], g['groupId'])):
                key = '%s|%s' % (g['orgunitId'], r['num'])
                if r['action'] in ('set', 'fix') and key not in seen:
                    seen.add(key)
                    nums.append(r['num'])
            if nums:
                plan.append((g, nums))
        total = sum(len(n) for g, n in plan)
        if total > AUTO_MAX:
            summary['stopped'] = ('%d names would change, more than the morning limit of %d. '
                                  'Nothing was changed. Run it from the page instead.' % (total, AUTO_MAX))
            entries.append({'when': summary['when'], 'by': 'Morning batch', 'source': 'Morning batch',
                            'group': '', 'num': '', 'before': '', 'after': '', 'pid': None,
                            'result': 'stopped', 'msg': summary['stopped']})
        else:
            for g, nums in plan:
                results, e = apply_group(g, nums, 'Morning batch', 'Morning batch')
                entries.extend(e)
                for x in results:
                    if x['result'] == 'changed':
                        summary['changed'] += 1
                    elif x['result'] == 'unconfirmed':
                        summary['unconfirmed'] += 1
                    elif x['result'] == 'error':
                        summary['errors'] += 1
    except Exception as e:
        summary['stopped'] = 'Error: ' + unicode(e)[:300]
        entries.append({'when': summary['when'], 'by': 'Morning batch', 'source': 'Morning batch',
                        'group': '', 'num': '', 'before': '', 'after': '', 'pid': None,
                        'result': 'error', 'msg': summary['stopped']})
    try:
        write_log(entries)
        settings['last_auto'] = summary
        save_settings(settings)
    except Exception:
        pass


# ---------------------------------------------------------------- delivery and usage
REPORT_PAGE = 1000        # rows per report request
REPORT_MAX_PAGES = 80     # ceiling per load; the page says when it was reached
SEND_GAP_MIN = 90         # rows of the same text further apart than this are separate sends
ECHO_HOURS = 12           # TouchPoint logs a text sent from PBT about 4 times as reports arrive
DUP_HOURS = 72            # same text to the same person again within this is a resend (waste); later is reuse
LATE_HOURS = 72           # a carrier can hold a text for a phone that is off for about 2 days before giving up
TYPE_LABEL = {
    'ScheduledBroadcast': 'Scheduled in PBT', 'Broadcast': 'Group text', 'BroadcastTest': 'Test send',
    # Unicast is a text to a list of numbers rather than to a whole PBT
    # group. TouchPoint sends every group text this way, with the numbers
    # it chose, so it is most of the volume, not one-to-one texting.
    'TwoWay': 'Reply', 'Twoway': 'Reply', 'Unicast': 'Sent to chosen people (how TouchPoint sends)', 'AutoResponse': 'Auto reply',
    'KeywordResponse': 'Keyword reply', 'Keyword': 'Keyword', 'ReservedKeyword': 'Keyword',
    'OptInResponse': 'Opt-in reply', 'StopResponse': 'Stop reply', 'StopAllResponse': 'Stop reply',
    'HelpResponse': 'Help reply', 'SystemMessage': 'System', 'Notification': 'Notification', 'Deact': 'Deactivated',
}
ALL_TYPES = list(TYPE_LABEL.keys())
GSM7X = set(u"@\u00a3$\u00a5\u00e8\u00e9\u00f9\u00ec\u00f2\u00c7\n\u00d8\u00f8\r\u00c5\u00e5\u0394_\u03a6\u0393\u039b\u03a9\u03a0\u03a8\u03a3\u0398\u039e"
           u"\u00c6\u00e6\u00df\u00c9 !\"#\u00a4%&'()*+,-./0123456789:;<=>?\u00a1ABCDEFGHIJKLMNOPQRSTUVWXYZ\u00c4\u00d6\u00d1\u00dc\u00a7"
           u"\u00bfabcdefghijklmnopqrstuvwxyz\u00e4\u00f6\u00f1\u00fc\u00e0^{}\\[~]|\u20ac")


def segments(text, mms):
    """Estimated billing segments for one recipient. Picture messages are
    billed as one MMS, so they count 1 here and are also counted separately."""
    if mms:
        return 1
    t = text or u''
    n = len(t)
    if n == 0:
        return 1
    if all(c in GSM7X for c in t):
        return 1 if n <= 160 else (n + 152) // 153
    return 1 if n <= 70 else (n + 66) // 67


def est_credits(text, mms):
    """Estimated PBT credits for one recipient, from PBT's help pages: a text
    is 1 credit per segment; PBT will not send a text over 3 segments, so a
    longer one goes as a text-only MMS; any MMS is 3 credits. Returns
    (credits, segments, special_chars, credits_if_plain, why)."""
    t = text or u''
    if mms:
        return 3, 0, False, 3, 'picture (MMS)'
    seg = segments(t, False)
    special = not all(c in GSM7X for c in t)
    plain_seg = segments(u''.join(c if c in GSM7X else u'?' for c in t), False)
    plain = plain_seg if plain_seg <= 3 else 3
    if seg > 3:
        return 3, seg, special, plain, 'over 3 segments, sent as text-only MMS'
    why = ''
    if seg > 1:
        why = '%d segments' % seg + (' (emoji or special characters)' if special and plain < seg else '')
    return seg, seg, special, plain, why


def parse_day(v, default):
    try:
        return datetime.datetime.strptime(v, '%Y-%m-%d')
    except Exception:
        return default


def _local_converter():
    """UTC -> church time for each moment, so daylight saving is right on
    every date. One fixed offset for the whole range put everything before a
    clock change an hour off (and made TouchPoint texts fail to match)."""
    name = model.Setting("LocalTimeZone", "Central Standard Time") or "Central Standard Time"
    try:
        import System
        zone = System.TimeZoneInfo.FindSystemTimeZoneById(name)

        def conv(dt):
            u = System.DateTime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second, System.DateTimeKind.Utc)
            l = System.TimeZoneInfo.ConvertTimeFromUtc(u, zone)
            return datetime.datetime(l.Year, l.Month, l.Day, l.Hour, l.Minute, l.Second)
        return name, conv
    except Exception:
        return None, None


def report_rows(group_ids, start, end, types):
    """Every report row in [start, end] (dates), all pages, each timestamp
    rewritten in church time. Asked for in UTC, a day wider on each side,
    then trimmed to the local dates. Returns (rows, notes)."""
    zone, conv = _local_converter()
    if conv:
        q_from, q_to = start - datetime.timedelta(days=1), end + datetime.timedelta(days=1)
    else:
        q_from, q_to = start, end
    body = {'pageNumber': 1, 'pageSize': REPORT_PAGE, 'messageTypes': types,
            'fromDate': q_from.strftime('%Y-%m-%d'), 'toDate': q_to.strftime('%Y-%m-%d')}
    if group_ids:
        body['groupIds'] = group_ids
    if conv:
        body['timeZoneOffset'] = '+00:00'
    else:
        zone, tz = church_offset()
        if tz:
            body['timeZoneOffset'] = tz
    rows, notes = [], ['Times are %s%s.' % (zone or 'PBT default time',
                                             ', daylight saving applied per date' if conv else '')]
    retried_day = False
    while True:
        res = pbt_post('/reports/message-reports', body)
        if res.get('isError'):
            code = res.get('appCode') or ''
            if code == 're.1006' and not retried_day:       # PBT is still on yesterday's date
                retried_day = True
                body['toDate'] = (end - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
                notes.append('PBT had not reached today yet, so today is not included.')
                continue
            if code == 're.1015':                            # no messages
                break
            raise Exception('PBT report: %s %s' % (code, res.get('message') or ''))
        data = res.get('data') or {}
        lo, hi = start, end + datetime.timedelta(days=1)
        for r in data.get('data') or []:
            if conv:
                t = _ts(r.get('timestamp'))
                if not t:
                    continue
                t = conv(t)
                if t < lo or t >= hi:
                    continue
                r['timestamp'] = t.strftime('%Y-%m-%d %H:%M:%S')
            rows.append(r)
        total = data.get('totalPages') or 1
        if body['pageNumber'] >= total:
            break
        if body['pageNumber'] >= REPORT_MAX_PAGES:
            notes.append('Stopped after %d rows (%d of %d pages). Pick a shorter range or one group.'
                         % (len(rows), body['pageNumber'], total))
            break
        body['pageNumber'] += 1
    return rows, notes


def _ts(v):
    v = (v or '').replace('T', ' ')[:19]
    try:
        return datetime.datetime.strptime(v, '%Y-%m-%d %H:%M:%S')
    except Exception:
        return None


def _status_kind(desc, code):
    d = (desc or code or '').lower()
    if 'deliver' in d and 'undeliver' not in d and 'not' not in d:
        return 'delivered'
    if any(k in d for k in ('fail', 'undeliver', 'reject', 'invalid', 'blocked', 'error', 'expired', 'stop')):
        return 'failed'
    return 'other'


def tp_sends_for(refs, group_ids_tp, start, end):
    """TouchPoint's own record of texts it handed to PBT: by reference id
    first, then every TouchPoint send to these groups in the window (for a
    text match when PBT does not carry the reference id)."""
    by_ref, lists = {}, {}
    refs = [r for r in refs if r and re.match(r'^[\w\-]{1,64}$', r)]
    for i in range(0, len(refs), 400):
        chunk = refs[i:i + 400]
        for r in q.QuerySql("""
                SELECT i.PbtReferenceId, i.ListID FROM SMSItems i WITH (NOLOCK)
                WHERE i.PbtReferenceId IN (%s)""" % ','.join("'%s'" % x for x in chunk)):
            by_ref[str(r.PbtReferenceId)] = int(r.ListID)
    where = "l.Created >= '%s' AND l.Created < '%s'" % (
        (start - datetime.timedelta(days=7)).strftime('%Y-%m-%d'), (end + datetime.timedelta(days=1)).strftime('%Y-%m-%d'))
    if group_ids_tp:
        where += " AND l.SendGroupID IN (%s)" % ','.join(str(int(x)) for x in group_ids_tp)
    for r in q.QuerySql("""
            SELECT l.ID, l.Created, l.SendAt, l.Message, l.SendGroupID, l.SentSMS, p.Name AS Sender
            FROM SMSList l WITH (NOLOCK)
            LEFT JOIN People p ON p.PeopleId = l.SenderID
            WHERE %s""" % where):
        lists[int(r.ID)] = {'id': int(r.ID), 'created': _ts(str(r.Created.ToString('yyyy-MM-dd HH:mm:ss'))),
                            'sendAt': _ts(str(r.SendAt.ToString('yyyy-MM-dd HH:mm:ss'))) if r.SendAt else None,
                            'message': norm(r.Message or ''), 'group': r.SendGroupID, 'n': r.SentSMS or 0,
                            'sender': r.Sender or ''}
    return by_ref, lists


def msg_key(m):
    return norm(re.sub(r'https?://\S+', '<link>', m or ''))


def send_key(r):
    return (r.get('groupId') or '', r.get('messageType') or '', msg_key(r.get('message')))


def split_runs(items, fresh=None):
    """items: [(time, row)] of one text to one group, sorted. Rows close
    together are one send. A small run up to LATE_HOURS after a bigger one is
    that send's late results: PBT stamps each row when the carrier answers,
    and for a phone that is off the carrier keeps trying for about two days
    before it reports back. fresh(t0, t1) says TouchPoint really sent the
    text again in that window, which keeps the run its own send.
    Returns [(run, late_rows)]."""
    runs = []
    for t, r in items:
        if runs and (t - runs[-1][-1][0]).total_seconds() <= SEND_GAP_MIN * 60:
            runs[-1].append((t, r))
        else:
            runs.append([(t, r)])
    out = []
    for run in runs:
        if out:
            main, late = out[-1]
            if (run[0][0] - main[0][0]).total_seconds() <= LATE_HOURS * 3600 \
                    and len(run) <= max(5, len(main) // 20) and len(main) >= 2 * len(run) \
                    and not (fresh and fresh(main[-1][0], run[0][0])):
                late.extend(run)
                continue
        out.append((run, []))
    return out


def send_runs(rows):
    """Outbound report rows -> [(key, run, late_rows)], every text and group."""
    sends = {}
    for r in rows:
        if not (r.get('direction') or '').lower().startswith('out'):
            continue
        t = _ts(r.get('timestamp'))
        if t:
            sends.setdefault(send_key(r), []).append((t, r))
    out = []
    for key, items in sends.items():
        items.sort(key=lambda x: x[0])
        for run, late in split_runs(items):
            out.append((key, run, late))
    return out


def tp_match_list(lists, gid_tp, key_text, first):
    """The TouchPoint send of this text due closest before `first` (within
    36 hours), for rows PBT reports without TouchPoint's reference id."""
    best = None
    for l in lists.values():
        if gid_tp and l['group'] != gid_tp:
            continue
        if not l['message'] or msg_key(l['message'])[:60] != key_text[:60]:
            continue
        due = l['sendAt'] or l['created']
        if not due or due > first + datetime.timedelta(minutes=5):
            continue
        gap = (first - due).total_seconds()
        if gap > 36 * 3600:
            continue
        if best is None or gap < (first - (best['sendAt'] or best['created'])).total_seconds():
            best = l
    return best


def trace(gid, start, end):
    groups = pbt_groups(include_system=True)
    by_ext = dict((g['groupId'], g) for g in groups)
    chosen = [g for g in groups if g['id'] == gid] if gid else []
    if gid and not chosen:
        raise Exception('Unknown group %d' % gid)
    types = [t for t in ALL_TYPES if t not in ('TwoWay', 'Twoway')]
    # PBT refuses the report without group ids (re.1010), so "all" is every group.
    rows, notes = report_rows([g['groupId'] for g in (chosen or groups)], start, end, types)
    sends = {}
    refs = set()
    for r in rows:
        if not (r.get('direction') or '').lower().startswith('out'):
            continue
        t = _ts(r.get('timestamp'))
        if not t:
            continue
        # Links are ignored: TouchPoint gives each person their own short
        # link, so otherwise one group text splits into a send per person.
        sends.setdefault(send_key(r), []).append((t, r))
        if r.get('referenceId'):
            refs.add(r.get('referenceId'))
    by_ref, lists = tp_sends_for(list(refs), [g['id'] for g in chosen], start, end)

    def person(t, r, stamp):
        return [digits10(r.get('subscriberNo')),
                ((r.get('subscriberFirstName') or '') + ' ' + (r.get('subscriberLastName') or '')).strip(),
                t.strftime(stamp), r.get('statusCodeDescription') or r.get('statusCode') or '',
                _status_kind(r.get('statusCodeDescription'), r.get('statusCode'))]
    out = []
    for key, items in sends.items():
        items.sort(key=lambda x: x[0])
        g = by_ext.get(key[0])
        gid_tp = g['id'] if g else None

        def fresh(t0, t1):
            # TouchPoint was asked to send this text again between the two runs.
            for l in lists.values():
                due = l['sendAt'] or l['created']
                if due and t0 < due <= t1 + datetime.timedelta(minutes=5) and (not gid_tp or l['group'] == gid_tp) \
                        and l['message'] and msg_key(l['message'])[:60] == key[2][:60]:
                    return True
            return False
        for run, late_rows in split_runs(items, fresh):
            first, last = run[0][0], run[-1][0]
            r0 = run[0][1]
            counts = {'delivered': 0, 'failed': 0, 'other': 0}
            people = []
            tp_list = None
            for t, r in run:
                p = person(t, r, '%H:%M:%S')
                counts[p[4]] += 1
                lid = by_ref.get(r.get('referenceId') or '')
                if lid and not tp_list:
                    tp_list = lists.get(lid) or {'id': lid}
                people.append(p)
            # Late results: a number already in the send is the carrier
            # reporting again; a new number is a recipient whose only answer
            # came late. Either way they are not part of how fast it went out.
            seen = set(p[0] for p in people)
            late = {'n': 0, 'again': 0, 'delivered': 0, 'failed': 0, 'other': 0, 'last': '', 'hours': 0}
            late_people = []
            for t, r in late_rows:
                p = person(t, r, '%m-%d %H:%M')
                late['n'] += 1
                late[p[4]] += 1
                if p[0] in seen:
                    late['again'] += 1
                else:
                    counts[p[4]] += 1
                    seen.add(p[0])
                late['last'] = t.strftime('%Y-%m-%d %H:%M:%S')
                late['hours'] = int((t - first).total_seconds() // 3600)
                late_people.append(p)
            match = 'reference'
            if not tp_list:
                match = 'text'
                tp_list = tp_match_list(lists, gid_tp, key[2], first)
            tp = None
            if tp_list and tp_list.get('created'):
                due = tp_list.get('sendAt') or tp_list['created']
                tp = {'listId': tp_list['id'], 'sender': tp_list.get('sender', ''),
                      'dueToLastMin': int((last - due).total_seconds() // 60),
                      'created': tp_list['created'].strftime('%Y-%m-%d %H:%M'),
                      'due': due.strftime('%Y-%m-%d %H:%M'),
                      'scheduled': bool(tp_list.get('sendAt') and (tp_list['sendAt'] - tp_list['created']).total_seconds() > 120),
                      'lateMin': int((first - due).total_seconds() // 60), 'match': match}
            offs = sorted(int((t - first).total_seconds()) for t, r in run)
            out.append({'first': first.strftime('%Y-%m-%d %H:%M:%S'), 'last': last.strftime('%Y-%m-%d %H:%M:%S'),
                        'spreadMin': int((last - first).total_seconds() // 60),
                        'avgSec': int(sum(offs) / max(1, len(offs))), 'medianSec': offs[len(offs) // 2] if offs else 0,
                        'group': g['name'] if g else (r0.get('groupName') or ''),
                        'type': r0.get('messageType') or '', 'typeLabel': TYPE_LABEL.get(r0.get('messageType') or '', r0.get('messageType') or ''),
                        'mms': bool(r0.get('mediaUrl')), 'message': r0.get('message') or '',
                        'n': len(seen), 'counts': counts, 'tp': tp, 'people': people,
                        'late': late if late['n'] else None, 'latePeople': late_people})
    # Name each number with the Names tab's owner rules (decide), so a
    # parent's phone on a child's record shows the parent. Each person row
    # becomes [num, name, time, status, kind, pid, why].
    nums = set()
    for o in out:
        for p in o['people'] + o['latePeople']:
            nums.add(p[0])
    found = tp_people_by_cell(list(nums))
    owner = {}
    for num in nums:
        first, last, pid, why, review = decide(found.get(num, []))
        owner[num] = (('%s %s' % (first, last)).strip() if pid else '', pid, why if (review or not pid) else '')
    for o in out:
        for p in o['people'] + o['latePeople']:
            name, pid, why = owner.get(p[0], ('', None, ''))
            pbt_name = p[1]
            p[1] = name or (pbt_name + ' (name in PBT)' if pbt_name else '')
            p.append(pid)
            p.append(why)
    out.sort(key=lambda x: x['first'], reverse=True)
    return {'sends': out, 'notes': notes, 'rowCount': len(rows)}


def tp_senders(groups, start, end, pbt_sent):
    """Who sent, from TouchPoint's own record (PBT's report has no sender).
    Only texts sent from TouchPoint have a name; the rest of PBT's total is
    one line, so the table still adds up."""
    ids = [int(g['id']) for g in groups]
    names = dict((int(g['id']), g['name']) for g in groups)
    if not ids:
        return []
    who = {}
    for r in q.QuerySql("""
            SELECT l.SenderID, p.Name AS Sender, l.SendGroupID, l.Message, ISNULL(l.SentSMS, 0) AS N
            FROM SMSList l WITH (NOLOCK)
            LEFT JOIN People p ON p.PeopleId = l.SenderID
            WHERE COALESCE(l.SendAt, l.Created) >= '%s' AND COALESCE(l.SendAt, l.Created) < '%s'
              AND l.SendGroupID IN (%s)""" % (start.strftime('%Y-%m-%d'),
                                                (end + datetime.timedelta(days=1)).strftime('%Y-%m-%d'),
                                                ','.join(str(i) for i in ids))):
        key = int(r.SenderID or 0)
        w = who.setdefault(key, {'sender': r.Sender or ('PeopleId %d' % key if key else 'No sender recorded'),
                                 'texts': 0, 'sent': 0, 'credits': 0, 'groups': {}})
        n = int(r.N or 0)
        w['texts'] += 1
        w['sent'] += n
        w['credits'] += n * est_credits(r.Message, False)[0]
        gname = names.get(int(r.SendGroupID or 0), '')
        w['groups'][gname] = w['groups'].get(gname, 0) + n
    out = []
    tp_sent = 0
    for w in who.values():
        tp_sent += w['sent']
        top = sorted(w['groups'].items(), key=lambda kv: -kv[1])
        w['groups'] = ', '.join('%s (%d)' % kv for kv in top[:3]) + (' +%d more' % (len(top) - 3) if len(top) > 3 else '')
        out.append(w)
    out.sort(key=lambda x: -x['credits'])
    rest = pbt_sent - tp_sent
    if rest > 0:
        out.append({'sender': 'Sent from PBT, or automatic replies (no name)', 'texts': None, 'sent': rest,
                    'credits': None, 'groups': ''})
    return out


def usage(gid, start, end):
    """Credits and where they go, from PBT's per-recipient report rows.
    Credits are estimates (see est_credits()); inbound texts are counted but not
    priced, since PBT does not say whether they cost credits."""
    groups = pbt_groups(include_system=True)
    by_ext = dict((g['groupId'], g) for g in groups)
    chosen = [g for g in groups if g['id'] == gid] if gid else []
    rows, notes = report_rows([g['groupId'] for g in (chosen or groups)], start, end, ALL_TYPES)
    tot = {'sent': 0, 'credits': 0, 'mms': 0, 'received': 0, 'failed': 0, 'failedCredits': 0,
           'fallbackCredits': 0, 'extraForSpecial': 0, 'longAsMms': 0, 'repeatCredits': 0}
    got = {}        # (number, group, text) -> last time it went to that person
    per_group, per_type, per_month, reached = {}, {}, {}, {}
    msgs, fails = {}, {}
    # A carrier's late answer and a picture's fallback text are not the
    # same text sent again, so neither counts as a repeat.
    late_ids = set(id(r) for key, run, late in send_runs(rows) for t, r in late)
    for r in rows:
        gname = (by_ext.get(r.get('groupId') or '') or {}).get('name') or r.get('groupName') or '(no group)'
        g = per_group.setdefault(gname, {'group': gname, 'sent': 0, 'credits': 0, 'mms': 0, 'received': 0, 'failed': 0,
                                         'failedCredits': 0, 'extra': 0, 'dupCredits': 0, 'fallbackCredits': 0})
        t = _ts(r.get('timestamp'))
        mon = t.strftime('%Y-%m') if t else '?'
        m = per_month.setdefault(mon, {'month': mon, 'sent': 0, 'credits': 0, 'mms': 0, 'received': 0})
        if (r.get('direction') or '').lower().startswith('in'):
            g['received'] += 1
            m['received'] += 1
            tot['received'] += 1
            continue
        mms = bool(r.get('mediaUrl'))
        cr, seg, special, plain, why = est_credits(r.get('message'), mms)
        failed = _status_kind(r.get('statusCodeDescription'), r.get('statusCode')) == 'failed'
        label = TYPE_LABEL.get(r.get('messageType') or '', r.get('messageType') or '')
        ty = per_type.setdefault(label, {'type': label, 'sent': 0, 'credits': 0})
        for b in (g, m, tot):
            b['sent'] += 1
            b['credits'] += cr
            b['mms'] += 1 if mms else 0
        ty['sent'] += 1
        ty['credits'] += cr
        num0 = digits10(r.get('subscriberNo'))
        dup_key = (num0, gname, msg_key(r.get('message')))
        prev = got.get(dup_key)
        countable = label != 'System' and not r.get('isFallback') and id(r) not in late_ids
        is_dup = bool(countable and prev and t and (t - prev).total_seconds() <= DUP_HOURS * 3600)
        if t and countable:
            got[dup_key] = t
        if is_dup:
            tot['repeatCredits'] += cr
            g['dupCredits'] += cr
        if not mms:
            g['extra'] += max(0, cr - plain)
        if failed:
            g['failed'] += 1
            g['failedCredits'] += cr
            tot['failed'] += 1
            tot['failedCredits'] += cr
            num = digits10(r.get('subscriberNo'))
            f = fails.setdefault(num, {'num': num, 'failed': 0, 'credits': 0, 'groups': set(), 'last': ''})
            f['failed'] += 1
            f['credits'] += cr
            f['groups'].add(gname)
            f['last'] = max(f['last'], r.get('timestamp') or '')
        if r.get('isFallback'):
            tot['fallbackCredits'] += cr
            g['fallbackCredits'] += cr
        tot['extraForSpecial'] += max(0, cr - plain) if not mms else 0
        if seg > 3:
            tot['longAsMms'] += 1
        reached.setdefault(gname, set()).add(digits10(r.get('subscriberNo')))
        if label.startswith('Sent to chosen') or label in ('Group text', 'Scheduled in PBT'):
            k = (gname, msg_key(r.get('message')))
            e = msgs.setdefault(k, {'group': gname, 'message': (r.get('message') or '')[:160], 'recipients': 0,
                                    'credits': 0, 'extra': 0, 'why': why, 'times': [], 'perPerson': cr,
                                    'dupPeople': 0, 'repeatCredits': 0})
            if is_dup:
                e['dupPeople'] += 1
                e['repeatCredits'] += cr
            e['recipients'] += 1
            e['credits'] += cr
            e['extra'] += max(0, cr - plain) if not mms else 0
            if t:
                e['times'].append(t)
    for gname, g in per_group.items():
        g['people'] = len(reached.get(gname, ()))
    # How many separate times each text went out (gaps longer than a send).
    for e in msgs.values():
        ts = sorted(e.pop('times'))
        runs = 1 if ts else 0
        for a, b in zip(ts, ts[1:]):
            if (b - a).total_seconds() > SEND_GAP_MIN * 60:
                runs += 1
        e['sends'] = runs
        e['first'] = ts[0].strftime('%m-%d %H:%M') if ts else ''
    costly = sorted(msgs.values(), key=lambda e: -e['credits'])[:15]
    # Only the same PERSON getting the same text again counts: a weekly
    # reminder that reaches different people each time is not waste.
    repeats = sorted([e for e in msgs.values() if e['dupPeople'] > 0], key=lambda e: -e['repeatCredits'])[:15]
    # Numbers that keep failing, named with the Names tab's owner rules.
    bad_all = [f for f in fails.values() if f['failed'] >= 3]
    tot['repeatFailCredits'] = sum(f['credits'] for f in bad_all)
    tot['repeatFailNumbers'] = len(bad_all)
    bad = sorted(bad_all, key=lambda f: -f['failed'])[:40]
    found = tp_people_by_cell([f['num'] for f in bad])
    for f in bad:
        first, last, pid, why, review = decide(found.get(f['num'], []))
        f['name'] = ('%s %s' % (first, last)).strip() if pid else ''
        f['pid'] = pid
        f['why'] = why if (review or not pid) else ''
        f['groups'] = ', '.join(sorted(f['groups']))
    senders = tp_senders(chosen or groups, start, end, tot['sent'])
    return {'total': tot, 'senders': senders, 'groups': sorted(per_group.values(), key=lambda x: -x['credits']),
            'types': sorted(per_type.values(), key=lambda x: -x['credits']),
            'months': sorted(per_month.values(), key=lambda x: x['month']),
            'costly': costly, 'repeats': repeats, 'failing': bad,
            'notes': notes, 'rowCount': len(rows)}


USAGE_VERSION = 6      # bump when the counting changes; older saved months are counted again


def tp_item_kind(status):
    """TouchPoint's per-recipient ResultStatus (PBT's delivery report as
    TouchPoint stored it) -> delivered, failed, nocredit or unknown."""
    st = (status or '').lower()
    if not st or 'queued' in st or 'sending' in st:
        return 'unknown'
    if 'insufficient credit' in st:
        return 'nocredit'
    if st == 'delivered':
        return 'delivered'
    return 'failed'


def _blank_counts():
    return {'credits': 0, 'sent': 0, 'failed': 0, 'received': 0, 'mms': 0, 'failedCredits': 0, 'extra': 0,
            'dupCredits': 0, 'fallbackCredits': 0, 'unknown': 0, 'noCredit': 0}


def tp_events(start, nxt, gname=None):
    """Every real text in [start, nxt) from TouchPoint's record, one per
    recipient, in time order, with the delivery result TouchPoint stored.
    TouchPoint logs a text sent from inside PBT about four times as PBT's
    reports come in, each as its own send, so a copy in ANOTHER send to the
    same number within ECHO_HOURS is that text again (best result kept).
    Two copies in ONE send are two people sharing a number: TouchPoint
    sends one per person, so that number really got it twice. A one-person
    copy 36 to LATE_HOURS later is the carrier's late answer for a phone
    that was off, which TouchPoint also logs as a new send."""
    names = dict((int(g['id']), g['name']) for g in pbt_groups(include_system=True))
    where = ''
    if gname:
        ids = [i for i, nm in names.items() if nm == gname]
        if not ids:
            return []
        where = ' AND l.SendGroupID IN (%s)' % ','.join(str(i) for i in ids)
    rank = {'delivered': 3, 'unknown': 2, 'failed': 1}
    events, open_, last, cache = [], {}, {}, {}
    for r in q.QuerySql("""
            SELECT l.ID, l.SendGroupID, l.Message, i.Number, i.ResultStatus, i.PeopleID, p.Name AS Sender,
                   COUNT(*) OVER (PARTITION BY i.ListID) AS ListSize,
                   CONVERT(varchar(19), COALESCE(l.SendAt, l.Created), 120) AS At
            FROM SMSItems i WITH (NOLOCK)
            JOIN SMSList l WITH (NOLOCK) ON l.ID = i.ListID
            LEFT JOIN People p ON p.PeopleId = l.SenderID
            WHERE COALESCE(l.SendAt, l.Created) >= '%s' AND COALESCE(l.SendAt, l.Created) < '%s'
              AND ISNULL(i.NoNumber, 0) = 0 AND ISNULL(i.NoOptIn, 0) = 0%s
            ORDER BY COALESCE(l.SendAt, l.Created), l.ID""" % (start.strftime('%Y-%m-%d'), nxt.strftime('%Y-%m-%d'), where)):
        gn = names.get(int(r.SendGroupID or 0))
        if not gn:
            continue
        lid = int(r.ID)
        if lid not in cache:
            cache[lid] = (est_credits(r.Message, False), msg_key(r.Message))
        (cr, seg, special, plain, why), key = cache[lid]
        kind = tp_item_kind(r.ResultStatus)
        t = _ts(r.At)
        num = re.sub(r'\D', '', r.Number or '')[-10:]
        ev = {'t': t, 'num': num, 'group': gn, 'key': key, 'text': r.Message or '', 'mms': False,
              'kind': kind, 'status': r.ResultStatus or '', 'fallback': False, 'late': False,
              'send': lid, 'sender': r.Sender or '', 'pid': r.PeopleID,
              'cr': cr, 'seg': seg, 'plain': plain, 'special': special, 'dupOf': None}
        if kind == 'nocredit':
            events.append(ev)
            continue
        dk = (num, gn, key)
        o = open_.get(dk)
        gap = (t - o['t']).total_seconds() if o and t and o['t'] else None
        if o and o['send'] != lid and gap is not None and (gap <= ECHO_HOURS * 3600 or (
                int(r.ListSize or 0) == 1 and 36 * 3600 <= gap <= LATE_HOURS * 3600)):
            if rank[kind] > rank[o['kind']]:
                o['kind'], o['status'] = kind, ev['status']
            continue
        prev = last.get(dk)
        if prev and t and prev['t'] and (t - prev['t']).total_seconds() <= DUP_HOURS * 3600:
            ev['dupOf'] = prev
        events.append(ev)
        open_[dk] = ev
        last[dk] = ev
    return events


def pbt_events(start, end, gname=None, rows=None):
    """The same list from PBT's report rows (periods PBT still reports).
    Late carrier answers and fallback texts are kept but never a repeat.
    Pass rows already read to skip reading them again."""
    groups = pbt_groups(include_system=True)
    by_ext = dict((g['groupId'], g) for g in groups)
    chosen = [g for g in groups if g['name'] == gname] if gname else groups
    if not chosen:
        return []
    if rows is None:
        rows, notes = report_rows([g['groupId'] for g in chosen], start, end, ALL_TYPES)
    refs = [r.get('referenceId') for r in rows if r.get('referenceId')]
    by_ref, lists = tp_sends_for(refs, [g['id'] for g in chosen] if gname else [], start, end)
    events = []
    for n, (key, run, late) in enumerate(send_runs(rows)):
        g = by_ext.get(key[0])
        gn = g['name'] if g else (run[0][1].get('groupName') or '(no group)')
        tp_list = None
        for t, r in run:
            lid = by_ref.get(r.get('referenceId') or '')
            if lid:
                tp_list = lists.get(lid)
                break
        if not tp_list:
            tp_list = tp_match_list(lists, g['id'] if g else None, key[2], run[0][0])
        label = TYPE_LABEL.get(key[1], key[1])
        sender = tp_list['sender'] if tp_list else ('Sent from PBT' if label != 'System' else 'PBT system message')
        for t, r in run + late:
            mms = bool(r.get('mediaUrl'))
            cr, seg, special, plain, why = est_credits(r.get('message'), mms)
            events.append({'t': t, 'num': digits10(r.get('subscriberNo')), 'group': gn, 'key': key[2],
                           'text': r.get('message') or '', 'mms': mms,
                           'kind': _status_kind(r.get('statusCodeDescription'), r.get('statusCode')),
                           'status': r.get('statusCodeDescription') or r.get('statusCode') or '',
                           'fallback': bool(r.get('isFallback')), 'late': (t, r) in late, 'system': label == 'System',
                           'send': n, 'sender': sender, 'pid': None,
                           'cr': cr, 'seg': seg, 'plain': plain, 'special': special, 'dupOf': None})
    events.sort(key=lambda e: e['t'])
    last = {}
    for e in events:
        if e['fallback'] or e['late'] or e['system']:
            continue
        dk = (e['num'], e['group'], e['key'])
        prev = last.get(dk)
        if prev and (e['t'] - prev['t']).total_seconds() <= DUP_HOURS * 3600:
            e['dupOf'] = prev
        last[dk] = e
    return events


def tp_month(month, start, nxt):
    """A period PBT no longer reports, counted from tp_events."""
    groups = {}
    tot = _blank_counts()
    tot['extraForSpecial'] = tot['repeatCredits'] = 0
    for e in tp_events(start, nxt):
        g = groups.setdefault(e['group'], _blank_counts())
        if e['kind'] == 'nocredit':
            g['noCredit'] += 1
            tot['noCredit'] += 1
            continue
        extra = max(0, e['cr'] - e['plain'])
        for b in (g, tot):
            b['credits'] += e['cr']
            b['sent'] += 1
        g['extra'] += extra
        tot['extraForSpecial'] += extra
        if e['kind'] == 'failed':
            for b in (g, tot):
                b['failed'] += 1
                b['failedCredits'] += e['cr']
        elif e['kind'] == 'unknown':
            g['unknown'] += 1
            tot['unknown'] += 1
        if e['dupOf']:
            g['dupCredits'] += e['cr']
            tot['repeatCredits'] += e['cr']
    return {'month': month, 'source': 'TouchPoint', 'v': USAGE_VERSION, 'total': tot, 'groups': groups}


# ---------------------------------------------------------------- click a number: what is behind it
CHAR_NAMES = {
    0x2018: ("curly apostrophe", "'"), 0x2019: ("curly apostrophe", "'"), 0x201C: ('curly quote', '"'),
    0x201D: ('curly quote', '"'), 0x2013: ('dash', '-'), 0x2014: ('long dash', '-'), 0x2026: ('three dots as one character', '...'),
    0x00A0: ('invisible no-break space', 'a normal space'), 0x200B: ('invisible zero-width space', 'delete it'),
    0x2022: ('bullet', '-'), 0xFE0F: ('invisible emoji style mark', 'goes with the emoji'),
    0x200D: ('invisible emoji joiner', 'goes with the emoji'),
}


def special_chars(text):
    """The characters that kept a text out of the plain text alphabet,
    each once with a count and what to type instead."""
    t = text or u''
    found, order = {}, []
    i = 0
    while i < len(t):
        c = t[i]
        o = ord(c)
        if 0xD800 <= o <= 0xDBFF and i + 1 < len(t):
            ch, name, fix = t[i:i + 2], 'emoji', 'leave it out, or accept the cost'
            i += 2
        else:
            i += 1
            if c in GSM7X:
                continue
            ch = c
            name, fix = CHAR_NAMES.get(o, (None, None))
            if not name:
                if o >= 0x2190:
                    name, fix = 'emoji or symbol', 'leave it out, or accept the cost'
                else:
                    name, fix = 'letter or mark outside the plain text alphabet (U+%04X)' % o, 'a plain letter'
        if ch not in found:
            found[ch] = {'ch': ch, 'name': name, 'fix': fix, 'n': 0}
            order.append(ch)
        found[ch]['n'] += 1
    return [found[c] for c in order]


def explain_text(text, mms):
    """One plain sentence on what this text costs per person and why."""
    cr, seg, special, plain, why = est_credits(text, mms)
    n = len(text or u'')
    if mms:
        return 'A picture message (MMS): 3 credits for each person.'
    chars = special_chars(text) if special else []
    names = ', '.join('%s %s%s' % (c['ch'], c['name'], ' x%d' % c['n'] if c['n'] > 1 else '') for c in chars)
    if cr > plain:
        s = ('Has %s. Any one of these sends the whole text as unicode, which fits 70 characters in a segment '
             'instead of 160. At %d characters it is %d segments' % (names, n, seg))
        s += (', and PBT sends anything over 3 segments as a 3 credit text-only MMS' if seg > 3 else '')
        s += '. Without them it would be %d credit%s, so it costs %d extra for each person.' % (
            plain, '' if plain == 1 else 's', cr - plain)
        return s
    if special:
        return 'Has %s, but at %d characters it still fits in %d segment%s, so no extra cost.' % (
            names, n, seg, '' if seg == 1 else 's')
    if seg > 1:
        return '%d characters: %d segments of plain text%s.' % (
            n, seg, ', sent as a 3 credit text-only MMS' if seg > 3 else '')
    return 'Fits in one plain text segment.'


DRILL_LIMIT = 400


def drill(month, gname, metric):
    start, nxt = period_bounds(month)
    today = datetime.datetime.now()
    end = min(nxt - datetime.timedelta(days=1), today)
    hist = load_json_content(USAGE_CONTENT, {})
    rec = hist.get(month) if isinstance(hist, dict) else None
    source = 'TouchPoint' if rec and rec.get('source') == 'TouchPoint' else 'PBT'
    if source == 'PBT':
        try:
            events = pbt_events(start, end, gname or None)
        except Exception as e:
            if 'ms.1003' not in unicode(e):
                raise
            source = 'TouchPoint'
    if source == 'TouchPoint':
        events = tp_events(start, nxt, gname or None)
    res = {'source': source, 'label': period_label(start, nxt), 'group': gname or 'All groups', 'metric': metric}
    sent = [e for e in events if e['kind'] != 'nocredit']
    if metric == 'dupCredits':
        res.update(drill_repeats([e for e in sent if e['dupOf']]))
    elif metric in ('failedCredits', 'unknown', 'noCredit', 'fallbackCredits'):
        pick = {'failedCredits': lambda e: e['kind'] == 'failed',
                'unknown': lambda e: e['kind'] == 'unknown',
                'noCredit': lambda e: e['kind'] == 'nocredit',
                'fallbackCredits': lambda e: e['fallback']}[metric]
        res.update(drill_list([e for e in events if pick(e)]))
    else:
        res.update(drill_messages(sent, metric))
    return res


def _names_for(nums):
    """Every active person in TouchPoint with each number, plus the owner
    the Names tab would pick."""
    found = tp_people_by_cell(list(nums))
    out = {}
    for num in nums:
        cands = found.get(num, [])
        first, last, pid, why, review = decide(cands)
        out[num] = {'owner': ('%s %s' % (first, last)).strip() if pid else '', 'pid': pid,
                    'all': ['%s %s%s' % (c['first'], c['last'], ' (%d)' % c['age'] if c['age'] is not None and c['age'] < 18 else '')
                            for c in cands]}
    return out


def _when(t):
    return t.strftime('%m-%d %H:%M') if t else ''


def _gap(sec):
    m = int(sec // 60)
    if m < 60:
        return '%d min' % m
    if m < 48 * 60:
        return '%dh %dm' % (m // 60, m % 60)
    return '%d days' % (m // 1440)


def drill_repeats(reps):
    names = _names_for(set(e['num'] for e in reps))
    causes = {}
    rows = []
    for e in reps:
        p = e['dupOf']
        nm = names.get(e['num']) or {'owner': '', 'pid': None, 'all': []}
        gap = (e['t'] - p['t']).total_seconds()
        if p['send'] == e['send']:
            cause = 'shared'
            who = nm['all']
            why = ('One send went to this number twice: TouchPoint sends one text per person, and %s share this cell phone.'
                   % (' and '.join(who) if len(who) <= 3 else '%d people' % len(who))) if len(who) > 1 else \
                'One send went to this number twice: two people on the send list have it as their cell phone.'
        elif p['kind'] == 'failed':
            cause = 'retry'
            why = 'Sent again %s later after the first one failed (%s).' % (_gap(gap), p['status'] or 'failed')
        elif p['sender'] and p['sender'] == e['sender']:
            cause = 'again'
            why = '%s sent the same text again %s later.' % (e['sender'], _gap(gap))
        else:
            cause = 'again'
            why = 'Sent again %s later: first by %s, then by %s.' % (_gap(gap), p['sender'] or 'unknown', e['sender'] or 'unknown')
        c = causes.setdefault(cause, {'cause': cause, 'texts': 0, 'credits': 0})
        c['texts'] += 1
        c['credits'] += e['cr']
        rows.append({'num': e['num'], 'owner': nm['owner'], 'pid': nm['pid'], 'group': e['group'],
                     'text': e['text'][:200], 'first': _when(p['t']), 'again': _when(e['t']), 'cause': cause,
                     'why': why, 'credits': e['cr'], 'status': e['status']})
    order = {'shared': 0, 'again': 1, 'retry': 2}
    rows.sort(key=lambda r: (order[r['cause']], r['again']))
    # The same send to many shared numbers is one finding, not dozens.
    sends = {}
    for e in reps:
        if e['dupOf']['send'] == e['send']:
            k = (e['group'], e['key'], e['send'])
            s = sends.setdefault(k, {'group': e['group'], 'text': e['text'][:200], 'when': _when(e['t']),
                                     'sender': e['sender'], 'numbers': 0, 'credits': 0})
            s['numbers'] += 1
            s['credits'] += e['cr']
    return {'kind': 'repeats', 'total': {'texts': len(reps), 'credits': sum(e['cr'] for e in reps)},
            'causes': sorted(causes.values(), key=lambda c: order[c['cause']]),
            'sharedSends': sorted(sends.values(), key=lambda s: -s['credits'])[:25],
            'rows': rows[:DRILL_LIMIT], 'more': max(0, len(rows) - DRILL_LIMIT)}


def drill_list(evs):
    names = _names_for(set(e['num'] for e in evs))
    st = {}
    for e in evs:
        k = e['status'] or 'No status'
        s = st.setdefault(k, {'status': k, 'texts': 0, 'credits': 0, 'numbers': set()})
        s['texts'] += 1
        s['credits'] += e['cr']
        s['numbers'].add(e['num'])
    for s in st.values():
        s['numbers'] = len(s['numbers'])
    per_num = {}
    for e in evs:
        per_num[e['num']] = per_num.get(e['num'], 0) + 1
    rows = [{'num': e['num'], 'owner': (names.get(e['num']) or {}).get('owner', ''), 'pid': (names.get(e['num']) or {}).get('pid'),
             'group': e['group'], 'text': e['text'][:200], 'when': _when(e['t']), 'status': e['status'],
             'sender': e['sender'], 'credits': e['cr'], 'timesThisNumber': per_num[e['num']],
             'late': e.get('late', False)} for e in evs]
    rows.sort(key=lambda r: (-r['timesThisNumber'], r['num'], r['when']))
    return {'kind': 'list', 'total': {'texts': len(evs), 'credits': sum(e['cr'] for e in evs),
                                      'numbers': len(per_num)},
            'statuses': sorted(st.values(), key=lambda s: -s['texts']),
            'rows': rows[:DRILL_LIMIT], 'more': max(0, len(rows) - DRILL_LIMIT)}


def drill_messages(evs, metric):
    """One row per text (all its sends in the period): who sent it, how many
    people, what each cost and why."""
    msgs = {}
    for e in evs:
        k = (e['group'], e['key'])
        m = msgs.get(k)
        if not m:
            m = msgs[k] = {'group': e['group'], 'text': e['text'], 'mms': e['mms'], 'first': e['t'], 'senders': {},
                           'sends': set(), 'numbers': set(), 'texts': 0, 'credits': 0, 'extra': 0,
                           'failedCredits': 0, 'dupCredits': 0, 'perPerson': e['cr'], 'plain': e['plain'], 'seg': e['seg']}
        m['senders'][e['sender'] or 'unknown'] = 1
        m['sends'].add(e['send'])
        m['numbers'].add(e['num'])
        m['texts'] += 1
        m['credits'] += e['cr']
        m['extra'] += 0 if e['mms'] else max(0, e['cr'] - e['plain'])
        m['failedCredits'] += e['cr'] if e['kind'] == 'failed' else 0
        m['dupCredits'] += e['cr'] if e['dupOf'] else 0
    rows = []
    for m in msgs.values():
        m['waste'] = m['extra'] + m['failedCredits'] + m['dupCredits']
        if metric == 'extra' and not m['extra']:
            continue
        if metric == 'waste' and not m['waste']:
            continue
        rows.append({'group': m['group'], 'text': m['text'][:400], 'first': _when(m['first']),
                     'senders': ', '.join(sorted(m['senders'])), 'sends': len(m['sends']),
                     'people': len(m['numbers']), 'texts': m['texts'], 'credits': m['credits'],
                     'perPerson': m['perPerson'], 'plain': m['plain'], 'seg': m['seg'], 'chars': len(m['text']),
                     'extra': m['extra'], 'failedCredits': m['failedCredits'], 'dupCredits': m['dupCredits'],
                     'waste': m['waste'], 'why': explain_text(m['text'], m['mms']),
                     'special': [] if m['mms'] else special_chars(m['text'])})
    sort_by = {'extra': 'extra', 'waste': 'waste', 'sent': 'texts'}.get(metric, 'credits')
    rows.sort(key=lambda r: -r[sort_by])
    tot = {}
    for k in ('texts', 'credits', 'extra', 'failedCredits', 'dupCredits', 'waste'):
        tot[k] = sum(r[k] for r in rows)
    return {'kind': 'messages', 'sortBy': sort_by, 'total': tot,
            'rows': rows[:DRILL_LIMIT], 'more': max(0, len(rows) - DRILL_LIMIT)}


def cycle_day():
    try:
        d = int(load_settings().get('cycle_day') or 1)
    except Exception:
        d = 1
    return max(1, min(d, 28))


def period_bounds(key, day=None):
    """'2026-10' with a cycle day of 18 -> (Oct 18, Nov 18): the billing
    period that STARTS in that month. Day 1 is a calendar month."""
    day = day or cycle_day()
    y, m = int(key[:4]), int(key[5:])
    start = datetime.datetime(y, m, day)
    nxt = datetime.datetime(y + (m // 12), m % 12 + 1, day)
    return start, nxt


def period_label(start, nxt):
    if start.day == 1:
        return start.strftime('%b %Y')
    last = nxt - datetime.timedelta(days=1)
    return '%s %d to %s %d' % (start.strftime('%b'), start.day, last.strftime('%b'), last.day)


def current_period_key(day=None):
    day = day or cycle_day()
    t = datetime.datetime.now()
    d = t if t.day >= day else (t.replace(day=1) - datetime.timedelta(days=1))
    return d.strftime('%Y-%m')


# ---------------------------------------------------------------- updates
def apply_update():
    """Admin only (checked by admin_post). Saves the newest TPxi Texting over
    this script, under whatever name it was saved as, so saved data (kept in
    Special Content under its own names) is untouched."""
    name = re.sub(r'[^\w.\- ]', '', _param('script_name') or '') or DC_SCRIPT_ID
    try:
        code = model.RestGet(DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID, {}) or ''
    except Exception as e:
        return {'success': False, 'message': 'Could not download the update: %s' % e}
    code = unicode(code)
    # Never save something that is not this script: an error page, an empty
    # reply or another tool would break the page for everyone.
    if len(code) < 5000 or 'TEXTING_VERSION' not in code or 'def app_page' not in code:
        return {'success': False, 'message': 'The download did not look like TPxi Texting, so nothing was changed.'}
    m = re.search(r"TEXTING_VERSION = '([\d.]+)'", code)
    got = m.group(1) if m else ''

    def parts(v):
        return [int(x) if x.isdigit() else 0 for x in (v or '0').split('.')]
    if parts(got) <= parts(TEXTING_VERSION):
        # The service lists a newer version than the file it serves. Saving it
        # would change nothing and the notice would keep coming back.
        return {'success': False, 'message': 'The published file still says version %s inside, so nothing was changed. '
                'Raise TEXTING_VERSION in the published copy to match the version listed.' % (got or 'unknown')}
    model.WriteContentPython(name, code)
    return {'success': True, 'version': m.group(1) if m else '', 'message': 'Updated %s. Reloading.' % name}


# ---------------------------------------------------------------- connection check
def pbt_check():
    """Each step between TouchPoint and PBT, in order, as
    {name, state: ok | warn | bad | info, detail}. Read only."""
    checks = []

    def add(name, state, detail):
        checks.append({'name': name, 'state': state, 'detail': detail})
    if not model.Setting('PBTClientID', '') or not model.Setting('PBTClientSecret', ''):
        add('PoweredByText keys', 'bad', 'PBTClientID or PBTClientSecret is empty in TouchPoint Settings. '
            'Nothing can reach PoweredByText until both are filled in.')
        return checks
    add('PoweredByText keys', 'ok', 'Both are filled in under TouchPoint Settings.')
    try:
        pbt_token()
        add('Sign in to PoweredByText', 'ok', 'Signed in at %s.' % _conn['api'].split('//')[-1])
    except Exception as e:
        add('Sign in to PoweredByText', 'bad', unicode(e) + ' Ask PoweredByText to confirm the keys.')
        return checks
    groups = pbt_groups(include_system=True)
    if not groups:
        add('Texting groups', 'bad', 'TouchPoint has no PoweredByText groups set up.')
        return checks
    found, errs = {}, []
    for o in sorted(set(g['orgunitId'] for g in groups)):
        page = 1
        while True:
            try:
                res = pbt_get('/orgunit/%s/groups?pageNumber=%d&pageSize=%d' % (o, page, PAGE_SIZE))
            except Exception as e:
                errs.append(unicode(e))
                break
            if res.get('isError'):
                errs.append('%s %s' % (res.get('appCode') or '', res.get('message') or ''))
                break
            data = res.get('data')
            rows = data.get('data') if isinstance(data, dict) else data
            for x in rows or []:
                if isinstance(x, dict):
                    found[str(x.get('groupId') or x.get('id') or '')] = 1
            pages = data.get('totalPages') if isinstance(data, dict) else 1
            if not rows or page >= (pages or 1):
                break
            page += 1
    missing = [g['name'] for g in groups if str(g['groupId']) not in found]
    if errs and not found:
        add('Texting groups', 'warn', 'Could not read the group list from PoweredByText: ' + errs[0])
    elif missing:
        add('Texting groups', 'bad', '%d of %d TouchPoint texting groups were not found in PoweredByText: %s. '
            'They may have been deleted or moved there.' % (len(missing), len(groups), ', '.join(missing[:10])))
    else:
        add('Texting groups', 'ok', 'All %d TouchPoint texting groups exist in PoweredByText.' % len(groups))
    try:
        y = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - datetime.timedelta(days=1)
        rows, notes = report_rows([g['groupId'] for g in groups], y, y, ALL_TYPES)
        add('Message report', 'ok', 'PoweredByText reported %d texts for yesterday. Delivery and Usage read this report.' % len(rows))
    except Exception as e:
        add('Message report', 'bad', 'PoweredByText would not give its message report: %s. Delivery and Usage need it.' % e)
    ids = ','.join(str(int(g['id'])) for g in groups)
    r = q.QuerySqlTop1("""
        SELECT CONVERT(varchar(16), MAX(DateReceived), 120) Last,
               SUM(CASE WHEN DateReceived >= DATEADD(day, -7, GETDATE()) THEN 1 ELSE 0 END) Week
        FROM SmsReceived WITH (NOLOCK) WHERE ToGroupId IN (%s)""" % ids)
    if r and r.Last:
        add('Replies reaching TouchPoint', 'ok' if (r.Week or 0) else 'warn',
            'The last text anyone sent to a group arrived %s; %d in the last 7 days.%s' % (
                r.Last, r.Week or 0, '' if (r.Week or 0) else ' If people have texted in this week, PoweredByText may not be passing replies to TouchPoint.'))
    else:
        add('Replies reaching TouchPoint', 'warn', 'TouchPoint has never received a text to these groups. '
            'Ask PoweredByText to check that replies are sent to TouchPoint.')
    r = q.QuerySqlTop1("""
        SELECT COUNT(*) N,
               SUM(CASE WHEN ISNULL(i.ResultStatus, '') = '' OR i.ResultStatus LIKE '%%queued%%'
                         OR i.ResultStatus LIKE '%%sending%%' THEN 1 ELSE 0 END) Waiting
        FROM SMSItems i WITH (NOLOCK) JOIN SMSList l WITH (NOLOCK) ON l.ID = i.ListID
        WHERE l.SendGroupID IN (%s) AND ISNULL(i.NoNumber, 0) = 0 AND ISNULL(i.NoOptIn, 0) = 0
          AND COALESCE(l.SendAt, l.Created) BETWEEN DATEADD(day, -3, GETDATE()) AND DATEADD(hour, -1, GETDATE())""" % ids)
    n, waiting = (r.N or 0, r.Waiting or 0) if r else (0, 0)
    if not n:
        add('Delivery results', 'info', 'TouchPoint sent no texts between 3 days and 1 hour ago, so there is nothing to check.')
    else:
        share = 100.0 * waiting / n
        add('Delivery results', 'ok' if share < 10 else 'warn' if share < 40 else 'bad',
            '%d of %d texts TouchPoint sent in the last 3 days (%.0f%%) are still waiting for a result after an hour.%s' % (
                waiting, n, share, '' if share < 10 else ' PoweredByText may not be sending delivery results back to TouchPoint.'))
    name, off = church_offset()
    add('Time zone', 'ok' if off else 'warn', 'TouchPoint is set to %s%s.' % (name, ' (%s right now)' % off if off else ', which could not be read'))
    st = load_settings()
    if not st.get('plan_credits') or not st.get('plan_price') or not st.get('cycle_day'):
        add('Plan and billing', 'warn', 'Not filled in under Settings, so plan use and dollars are not shown.')
    else:
        add('Plan and billing', 'ok', '%s credits for $%g, billing period starts on day %s.' % (
            '{:,}'.format(st['plan_credits']), st['plan_price'], st['cycle_day']))
    return checks


def pbt_check_links():
    """Asks PBT to shorten one link (it adds that link to the first group's
    short link list in PoweredByText)."""
    g = ([x for x in pbt_groups()] or [None])[0]
    if not g:
        raise Exception('TouchPoint has no PoweredByText groups set up.')
    res = pbt_post('/orgunit/%s/groups/%s/shortUrls' % (g['orgunitId'], g['groupId']), {'longUrl': 'https://www.poweredbytext.com/'})
    url = ((res or {}).get('data') or {}).get('url') if not (res or {}).get('isError') else None
    if url:
        return {'state': 'ok', 'detail': 'Short links work: PoweredByText gave back %s (in %s).' % (url, g['name'])}
    return {'state': 'bad', 'detail': 'PoweredByText would not shorten a link: %s %s. Turn on M-Click (short links) in '
            'PoweredByText settings; until then links go out as typed.' % ((res or {}).get('appCode') or '', (res or {}).get('message') or '')}


# ---------------------------------------------------------------- usage saved by day
# Each refresh reads PBT only from the last good refresh (less a few days, so
# a late carrier answer or a repeat that crosses the line is still counted
# right) through today, and replaces just those days. Older days stay saved,
# so a billing period is the sum of its days and nothing is read twice.
# Counts only: no numbers, names or text are kept.
USAGE_DAYS_CONTENT = 'TPxi_PBTNameSyncUsageDays'
DAYS_VERSION = 1
DAY_KEYS = ('credits', 'sent', 'mms', 'received', 'failed', 'failedCredits', 'extra', 'dupCredits', 'fallbackCredits')
REFRESH_OVERLAP_DAYS = 3
PBT_KEEP_DAYS = 58        # PBT reports about two months back
DAYS_KEEP = 800           # about two years of days


def load_days():
    d = load_json_content(USAGE_DAYS_CONTENT, {})
    if not isinstance(d, dict) or d.get('v') != DAYS_VERSION or not isinstance(d.get('days'), dict):
        d = {'v': DAYS_VERSION, 'days': {}, 'from': '', 'through': ''}
    return d


REFRESH_SLICE_DAYS = 7     # the page reads a week at a time so it can show progress


def refresh_plan():
    """Where the next refresh starts, split into week-long pieces."""
    store = load_days()
    today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    floor = today - datetime.timedelta(days=PBT_KEEP_DAYS)
    last = parse_day((store.get('through') or '')[:10], None)
    start = max(floor, last - datetime.timedelta(days=REFRESH_OVERLAP_DAYS)) if last else floor
    slices, a = [], start
    while a <= today:
        b = min(a + datetime.timedelta(days=REFRESH_SLICE_DAYS - 1), today)
        slices.append([a.strftime('%Y-%m-%d'), b.strftime('%Y-%m-%d')])
        a = b + datetime.timedelta(days=1)
    return {'success': True, 'slices': slices, 'through': store.get('through') or ''}


def usage_refresh(sfrom=None, sto=None, first_piece=True, last_piece=True):
    """Reads [sfrom, sto] (default: from the last refresh through today) and
    replaces those days. The first piece settles where the saved days are
    complete from; the last one records the refresh time."""
    store = load_days()
    today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    floor = today - datetime.timedelta(days=PBT_KEEP_DAYS)
    last = parse_day((store.get('through') or '')[:10], None)
    start = max(floor, last - datetime.timedelta(days=REFRESH_OVERLAP_DAYS)) if last else floor
    end = today
    if sfrom:
        start = max(floor, parse_day(sfrom, start))
        end = min(today, parse_day(sto, today))
    ctx = max(floor, start - datetime.timedelta(days=REFRESH_OVERLAP_DAYS))
    groups = pbt_groups(include_system=True)
    by_ext = dict((g['groupId'], g) for g in groups)
    rows, notes = report_rows([g['groupId'] for g in groups], ctx, end, ALL_TYPES)
    first = start.strftime('%Y-%m-%d')
    stop = end.strftime('%Y-%m-%d')
    fresh = {}

    def bucket(day, gname):
        d = fresh.setdefault(day, {})
        return d.setdefault(gname, dict((k, 0) for k in DAY_KEYS))
    for e in pbt_events(ctx, today, rows=rows):
        day = e['t'].strftime('%Y-%m-%d')
        if day < first or day > stop:
            continue
        b = bucket(day, e['group'])
        b['credits'] += e['cr']
        b['sent'] += 1
        b['mms'] += 1 if e['mms'] else 0
        b['extra'] += 0 if e['mms'] else max(0, e['cr'] - e['plain'])
        if e['kind'] == 'failed':
            b['failed'] += 1
            b['failedCredits'] += e['cr']
        if e['dupOf']:
            b['dupCredits'] += e['cr']
        if e['fallback']:
            b['fallbackCredits'] += e['cr']
    for r in rows:
        if (r.get('direction') or '').lower().startswith('in'):
            t = _ts(r.get('timestamp'))
            if t and first <= t.strftime('%Y-%m-%d') <= stop:
                g = by_ext.get(r.get('groupId') or '')
                bucket(t.strftime('%Y-%m-%d'), (g or {}).get('name') or r.get('groupName') or '(no group)')['received'] += 1
    days = store['days']
    for day in [d for d in days if first <= d <= stop]:
        del days[day]
    for day, gs in fresh.items():
        days[day] = dict((gname, dict((k, v) for k, v in c.items() if v)) for gname, c in gs.items())
    cutoff = (today - datetime.timedelta(days=DAYS_KEEP)).strftime('%Y-%m-%d')
    for day in [d for d in days if d < cutoff]:
        del days[day]
    # Saved days are complete from 'from' on. A refresh after more than PBT's
    # two months leaves a hole, so the complete stretch restarts here.
    if first_piece and (not last or last - datetime.timedelta(days=REFRESH_OVERLAP_DAYS) < floor or not store.get('from')):
        store['from'] = first
    store['from'] = max(store.get('from') or first, cutoff)
    if last_piece:
        store['through'] = now_str()
    model.WriteContentText(USAGE_DAYS_CONTENT, safe_json(store), "")
    return {'from': store['from'], 'through': store['through'], 'rows': len(rows), 'read': first, 'notes': notes}


def period_from_days(month, start, nxt, store):
    """A billing period added up from saved days, or None when the saved
    days do not reach back to its start."""
    if not store.get('from') or start.strftime('%Y-%m-%d') < store['from']:
        return None
    s0, n0 = start.strftime('%Y-%m-%d'), nxt.strftime('%Y-%m-%d')
    groups = {}
    total = dict((k, 0) for k in DAY_KEYS)
    for day, gs in store['days'].items():
        if not (s0 <= day < n0):
            continue
        for gname, c in gs.items():
            g = groups.setdefault(gname, dict((k, 0) for k in DAY_KEYS))
            for k, v in c.items():
                if k in g:
                    g[k] += v
                    total[k] += v
    t = dict(total)
    t['extraForSpecial'] = t.pop('extra')
    t['repeatCredits'] = t.pop('dupCredits')
    return {'month': month, 'counted': store.get('through') or now_str(), 'v': USAGE_VERSION, 'source': 'PBT',
            'byDay': True, 'total': t, 'groups': groups}


def save_last_month():
    """Morning batch: save last month while PBT still reports it (it keeps
    about two months), so history builds without anyone opening the page."""
    try:
        cs, cn = period_bounds(current_period_key())
        month = (cs - datetime.timedelta(days=1)).strftime('%Y-%m')
        hist = load_json_content(USAGE_CONTENT, {})
        cur = hist.get(month) if isinstance(hist, dict) else None
        if not cur or cur.get('v') != USAGE_VERSION or cur.get('source') != 'PBT' \
                or cur.get('cycleDay', 1) != cycle_day():
            usage_month(month)
    except Exception:
        pass


def usage_month(month):
    """Credits for one calendar month, every group. A finished month cannot
    change, so it is saved and never read from PBT again; the current month
    is always counted fresh."""
    if not re.match(r'^\d{4}-\d{2}$', month or ''):
        raise Exception('Month must look like 2026-10')
    start, nxt = period_bounds(month)
    today = datetime.datetime.now()
    if start > today:
        raise Exception('That billing period has not started.')
    end = min(nxt - datetime.timedelta(days=1), today)
    complete = nxt <= today.replace(hour=0, minute=0, second=0, microsecond=0)
    store = load_days()
    rec = period_from_days(month, start, nxt, store)
    if rec is not None:
        # Finished only once a refresh has run after the period ended.
        complete = complete and (store.get('through') or '')[:10] >= nxt.strftime('%Y-%m-%d')
        add_tp_compare(rec, month, start, nxt)
        return save_month(rec, complete)
    try:
        u = usage(0, start, end)
    except Exception as e:
        # PBT keeps about two months of report data (ms.1003). Older months
        # come from TouchPoint's own record instead: texts it sent only.
        if 'ms.1003' not in unicode(e):
            raise
        return save_month(tp_month(month, start, nxt), complete)
    t = u['total']
    rec = {'month': month, 'counted': now_str(), 'complete': complete, 'v': USAGE_VERSION, 'source': 'PBT',
           'total': dict((k, t.get(k, 0)) for k in ('credits', 'sent', 'mms', 'received', 'failed', 'failedCredits',
                                                    'extraForSpecial', 'repeatCredits', 'fallbackCredits')),
           'groups': dict((g['group'], {'credits': g['credits'], 'sent': g['sent'], 'failed': g['failed'],
                                        'received': g['received'], 'mms': g['mms'], 'failedCredits': g['failedCredits'],
                                        'extra': g['extra'], 'dupCredits': g['dupCredits'],
                                        'fallbackCredits': g['fallbackCredits']}) for g in u['groups'])}
    add_tp_compare(rec, month, start, nxt)
    return save_month(rec, complete)


def add_tp_compare(rec, month, start, nxt):
    """TouchPoint's own record for the same period, to show what went around it."""
    try:
        tp = tp_month(month, start, nxt)
        rec['tpTotal'] = {'credits': tp['total']['credits'], 'sent': tp['total']['sent']}
        for gname, g in rec['groups'].items():
            tg = tp['groups'].get(gname) or {}
            g['tpCredits'], g['tpSent'] = tg.get('credits', 0), tg.get('sent', 0)
        for gname, tg in tp['groups'].items():
            if gname not in rec['groups']:
                rec['groups'][gname] = {'credits': 0, 'sent': 0, 'tpCredits': tg['credits'], 'tpSent': tg['sent']}
    except Exception:
        pass


def save_month(rec, complete):
    rec['complete'] = complete
    rec['cycleDay'] = cycle_day()
    s0, n0 = period_bounds(rec['month'])
    rec['label'] = period_label(s0, n0)
    rec.setdefault('v', USAGE_VERSION)
    rec.setdefault('counted', now_str())
    if complete:
        hist = load_json_content(USAGE_CONTENT, {})
        if not isinstance(hist, dict):
            hist = {}
        hist[rec['month']] = rec
        try:
            model.WriteContentText(USAGE_CONTENT, safe_json(hist), "")
        except Exception:
            pass
    return rec



# ================================================================ pages and actions
def inbox_post():
    action = _param('action')
    try:
        if not ME:
            raise Exception('Please sign in.')
        if not IS_STAFF:
            raise Exception('The texting inbox is for staff (Access role).')
        if action == 'threads':
            res = act_threads()
        elif action == 'thread':
            res = act_thread()
        elif action == 'send':
            res = act_send()
        elif action == 'done':
            res = act_done()
        elif action == 'mark_read':
            res = act_mark_read()
        elif action == 'find':
            res = act_find()
        elif action == 'group_info':
            res = act_group_info()
        elif action == 'upload_media':
            res = act_upload_media()
        elif action == 'group_send':
            res = act_group_send()
        elif action == 'people':
            res = act_people()
        elif action == 'pbt_test':
            res = act_pbt_test()
        elif action == 'pbt_speed':
            res = act_pbt_speed()
        else:
            res = {'success': False, 'message': 'Unknown action'}
        print safe_json(res)
    except Exception as e:
        print safe_json({'success': False, 'message': unicode(e)})


def admin_post():
    action = _param('action')
    try:
        if not model.UserIsInRole("Admin"):
            raise Exception("Admin only.")
        if action in ('preview_group', 'apply'):
            gid = int(_param('gid', '0'))
            g = [x for x in pbt_groups() if x['id'] == gid]
            if not g:
                raise Exception("Unknown group %d" % gid)
            g = g[0]
        if action == 'preview_group':
            rows = evaluate(g, pbt_subscribers(g['orgunitId'], g['groupId']))
            print safe_json({'success': True, 'group': g['name'], 'orgunitId': g['orgunitId'], 'rows': rows})
        elif action == 'apply':
            if not NAME_WRITES:
                raise Exception('Changing names in PoweredByText is on hold until PoweredByText adds a way to do it.')
            nums = [n for n in _param('nums').split(',') if re.match(r'^\d{10}$', n)]
            if not nums:
                raise Exception("No numbers given.")
            who = model.UserName or 'Unknown'
            results, entries = apply_group(g, nums, 'Page', who)
            write_log(entries)
            print safe_json({'success': True, 'results': results})
        elif action == 'usage_history':
            hist = load_json_content(USAGE_CONTENT, {})
            print safe_json({'success': True, 'months': hist if isinstance(hist, dict) else {},
                             'creditCost': load_settings().get('credit_cost'), 'cycleDay': cycle_day(),
                             'current': current_period_key(), 'version': USAGE_VERSION,
                             'daysFrom': load_days().get('from'), 'daysThrough': load_days().get('through')})
        elif action == 'usage_refresh_plan':
            print safe_json(refresh_plan())
        elif action == 'usage_refresh':
            r = usage_refresh(_param('sfrom') or None, _param('sto') or None,
                              _param('firstpiece', '1') == '1', _param('lastpiece', '1') == '1')
            r['success'] = True
            print safe_json(r)
        elif action == 'usage_month':
            print safe_json({'success': True, 'rec': usage_month(_param('month'))})
        elif action == 'apply_update':
            print safe_json(apply_update())
        elif action == 'pbt_check':
            print safe_json({'success': True, 'checks': pbt_check()})
        elif action == 'pbt_check_links':
            r = pbt_check_links()
            r['success'] = True
            print safe_json(r)
        elif action == 'drill':
            month = _param('month')
            if not re.match(r'^\d{4}-\d{2}$', month):
                raise Exception('Month must look like 2026-10')
            r = drill(month, _param('grp'), _param('metric') or 'credits')
            r['success'] = True
            print safe_json(r)
        elif action == 'save_cost':
            st = load_settings()
            v = _param('cost')
            try:
                st['credit_cost'] = round(float(v), 5) if v else None
            except Exception:
                raise Exception('Cost per credit must be a number, such as 0.035')
            for key, label in (('plan_credits', 'Credits in the plan'), ('plan_price', 'Plan price'), ('overage_cost', 'Overage per credit')):
                v = _param(key)
                try:
                    st[key] = (int(float(v)) if key == 'plan_credits' else round(float(v), 5)) if v else None
                except Exception:
                    raise Exception('%s must be a number.' % label)
            if _param('admin_all') in ('0', '1'):
                st['admin_all_groups'] = _param('admin_all') == '1'
            if _param('emoji_policy') in EMOJI_POLICIES:
                st['emoji_policy'] = _param('emoji_policy')
            if _param('attach_policy') in ATTACH_POLICIES:
                st['attach_policy'] = _param('attach_policy')
            day = _param('cycle_day')
            if day:
                if not day.isdigit() or not 1 <= int(day) <= 28:
                    raise Exception('The billing cycle start day must be 1 to 28.')
                st['cycle_day'] = int(day)
            save_settings(st)
            print safe_json({'success': True, 'cost': st['credit_cost'], 'cycleDay': st.get('cycle_day') or 1,
                             'planCredits': st.get('plan_credits'), 'planPrice': st.get('plan_price'),
                             'overageCost': st.get('overage_cost')})
        elif action in ('trace', 'usage'):
            today = datetime.datetime.now()
            start = parse_day(_param('from'), today - datetime.timedelta(days=6))
            end = parse_day(_param('to'), today)
            if end < start:
                start, end = end, start
            if (end - start).days > 61:
                raise Exception('PBT reports cover at most 2 months at a time.')
            gid = int(_param('gid', '0') or 0)
            res = trace(gid, start, end) if action == 'trace' else usage(gid, start, end)
            res['success'] = True
            res['creditCost'] = load_settings().get('credit_cost')
            print safe_json(res)
        elif action == 'log':
            log = load_json_content(LOG_CONTENT, [])
            print safe_json({'success': True, 'log': log if isinstance(log, list) else []})
        elif action == 'save_settings':
            s = load_settings()
            s['auto'] = _param('auto') == '1' and NAME_WRITES
            s['groups'] = [int(x) for x in _param('groups').split(',') if x.isdigit()]
            name = _param('script_name')
            if s['auto']:
                if not re.match(r'^[\w\-. ]+$', name or ''):
                    raise Exception("Could not tell this script's name from the address bar.")
                mb_install(name)
            else:
                mb_uninstall()
            save_settings(s)
            print safe_json({'success': True, 'installed': MB_START in mb_read()})
        else:
            print safe_json({'success': False, 'message': 'Unknown action'})
    except Exception as e:
        print safe_json({'success': False, 'message': unicode(e)})


def inbox_page():
    me = my_name()
    groups = my_groups() if ME and IS_STAFF else {}
    if ME and not IS_STAFF:
        return ('<div class="txmsg">Texting is for staff (Access role).</div>', False)
    elif not groups:
        return ('<div class="txmsg">You are not a member of any PoweredByText texting group in TouchPoint. '
                'Ask an administrator to add you to one.</div>', False)
    else:
        first = (me.First if me else '') or ''
        can_new = any(can_send_in(g) for g in groups.values())
        css = """
<style>
#tx{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#1f2937;--line:#e5e7eb;--muted:#6b7280;--brand:#1e3a5f}
#tx *{box-sizing:border-box}
#tx .wrap{display:flex;border:1px solid var(--line);border-radius:10px;height:calc(100vh - 170px);min-height:520px;background:#fff;overflow:hidden}
#tx .left{width:360px;flex:none;min-width:0;border-right:1px solid var(--line);display:flex;flex-direction:column}
#tx .right{flex:1;display:flex;flex-direction:column;min-width:0}
#tx .tools{padding:10px;border-bottom:1px solid var(--line);display:flex;flex-wrap:wrap;gap:6px}
#tx .tools select,#tx .tools input{padding:6px 8px;border:1px solid #d1d5db;border-radius:6px;font-size:13px}
#tx .tools input{flex:1;min-width:120px}
#tx .seg{display:flex;border:1px solid #d1d5db;border-radius:6px;overflow:hidden}
#tx .seg a{padding:5px 10px;font-size:12px;color:#374151;text-decoration:none}
#tx .seg a.on{background:var(--brand);color:#fff}
#tx .list{overflow-y:auto;flex:1}
#tx .th{padding:10px 12px;border-bottom:1px solid #f3f4f6;cursor:pointer;position:relative}
#tx .th:hover{background:#f9fafb}
#tx .th.on{background:#eef2f7}
#tx .th .top{display:flex;justify-content:space-between;gap:8px}
#tx .th .nm{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#tx .th .tm{font-size:11px;color:var(--muted);white-space:nowrap}
#tx .th .pv{font-size:13px;color:#4b5563;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:2px}
#tx .th .meta{font-size:11px;color:var(--muted);margin-top:3px;display:flex;gap:6px;flex-wrap:wrap}
#tx .th.isnew .nm::before{content:'';display:inline-block;width:8px;height:8px;border-radius:50%;background:#2563eb;margin-right:6px;vertical-align:middle}
#tx .pill{display:inline-block;padding:0 6px;border-radius:9px;font-size:11px;font-weight:600}
#tx .p-new{background:#dbeafe;color:#1e40af}
#tx .th.isnew .nm{font-weight:800}
#tx .p-notanswered{background:#fef3c7;color:#92400e}
#tx .p-done{background:#e5e7eb;color:#374151}
#tx .p-optedout{background:#fee2e2;color:#991b1b}
#tx .p-answered{background:#dcfce7;color:#166534}
#tx .p-other{background:#f3f4f6;color:#6b7280}
#tx .head{padding:12px 16px;border-bottom:1px solid var(--line)}
#tx .head h3{margin:0;font-size:17px}
#tx .head .sub{font-size:12px;color:var(--muted);margin-top:2px}
#tx .head .ppl{font-size:12px;margin-top:6px}
#tx .head .ppl a{margin-right:10px}
#tx .head .acts{float:right;display:flex;gap:6px}
#tx .msgs{flex:1;overflow-y:auto;padding:14px 16px;background:#f9fafb}
#tx .day{text-align:center;font-size:11px;color:var(--muted);margin:10px 0}
#tx .m{max-width:72%;margin:4px 0;display:flex;flex-direction:column}
#tx .m .b{padding:8px 12px;border-radius:14px;white-space:pre-wrap;word-wrap:break-word;font-size:14px;line-height:1.35}
#tx .m .i{font-size:11px;color:var(--muted);margin:2px 6px}
#tx .m.in{align-items:flex-start}
#tx .m.in .b{background:#fff;border:1px solid var(--line)}
#tx .m.out{margin-left:auto;align-items:flex-end}
#tx .m.out .b{background:var(--brand);color:#fff}
#tx .m.bc .b{background:#e5e7eb;color:#374151}
#tx .m.auto .b{background:#f3f4f6;color:#6b7280;border:1px dashed #d1d5db;font-size:12px}
#tx .m.rx .b{background:transparent;border:1px dashed #d1d5db;color:var(--muted);font-size:12px}
#tx .m.stop .b{background:#fee2e2;color:#991b1b;border:none}
#tx .compose{border-top:1px solid var(--line);padding:10px 14px}
#tx .compose textarea{width:100%;min-height:64px;max-height:200px;padding:8px 10px;border:1px solid #d1d5db;border-radius:8px;font:inherit;resize:vertical}
#tx .compose .crow{display:flex;gap:8px;align-items:flex-end}
#tx .compose .crow textarea{flex:1;min-width:0;min-height:42px;resize:none;line-height:1.35}
#tx .compose .crow button{flex:none;height:42px;padding:0 16px}
#tx .pplbtn{display:none;font-size:12px}
#tx .compose .row{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-top:6px;flex-wrap:wrap;font-size:12px;color:var(--muted)}
#tx label.sig{display:inline-flex;align-items:center;gap:6px;font-weight:normal;margin:0;cursor:pointer;font-size:12px;color:var(--muted)}
#tx label.sig input[type=checkbox]{display:inline-block;width:auto;height:auto;margin:0;position:static}
#tx .note{padding:10px 14px;font-size:13px;border-top:1px solid var(--line);background:#fffbeb;color:#92400e}
#tx .note.red{background:#fef2f2;color:#991b1b}
#tx button{padding:6px 12px;border-radius:6px;border:1px solid var(--brand);background:var(--brand);color:#fff;cursor:pointer;font-size:13px}
#tx button.ghost{background:#fff;color:var(--brand)}
#tx button:disabled{opacity:.5;cursor:default}
#tx .empty{color:var(--muted);text-align:center;margin-top:80px}
#tx .kid{color:#991b1b}
#tx .find{position:absolute;inset:0;background:rgba(17,24,39,.35);display:none;align-items:flex-start;justify-content:center;padding-top:80px;z-index:50}
#tx .find .card{background:#fff;border-radius:10px;width:min(560px,94vw);padding:16px;max-height:70vh;overflow:auto}
#tx .find input{width:100%;padding:8px 10px;border:1px solid #d1d5db;border-radius:6px}
#tx .find .pr{padding:8px 0;border-bottom:1px solid #f3f4f6;font-size:13px}
#tx .find .pr button{margin:4px 6px 0 0;padding:3px 8px;font-size:12px}
#tx .back{display:none}
#tx .qr{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:6px}
#tx .qr button{padding:3px 10px;font-size:12px;border-radius:14px;background:#fff;color:var(--brand)}
#tx .meter{font-size:12px;color:var(--muted);margin-top:6px}
#tx .xtools{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-top:6px;font-size:12px;color:var(--muted)}
#tx .xtools .xbtn{padding:3px 10px;font-size:12px}
#tx .xchip{display:inline-flex;gap:6px;align-items:center;border:1px solid var(--line);border-radius:14px;padding:2px 4px 2px 10px;background:#f9fafb;color:#1f2937}
#tx .xchip .xx{border:0;background:transparent;color:var(--muted);padding:0 6px;font-size:14px}
#tx .xfrom{margin-left:auto}
#tx .xfrom select{padding:2px 4px;font-size:12px;border:1px solid #d1d5db;border-radius:6px}
#tx .xpick{flex-wrap:wrap;gap:2px;margin-top:6px;padding:6px;border:1px solid var(--line);border-radius:8px;background:#fff}
#tx .xpick button{border:0;background:transparent;color:inherit;font-size:20px;padding:2px 4px}
#tx .xline{font-size:12px;color:var(--muted);margin-top:4px}
#tx .xbad{color:#991b1b}
#tx .xwarn{color:#92400e}
#tx .xfb textarea{width:100%;min-height:44px;margin-top:6px;padding:6px 8px;border:1px solid #d1d5db;border-radius:8px;font:inherit}
#tx .meter b{color:#1f2937}
#tx .meter .warn{margin-top:4px;padding:6px 8px;border-radius:6px;background:#fffbeb;color:#92400e}
#tx .meter .warn button{margin-left:6px;padding:2px 8px;font-size:12px}
#tx .meter .ch{display:inline-block;min-width:16px;padding:0 3px;margin:0 1px;border:1px solid #f59e0b;border-radius:4px;background:#fff;color:#1f2937;text-align:center}
#tx #gmodal textarea{width:100%;min-height:110px;padding:8px 10px;border:1px solid #d1d5db;border-radius:8px;font:inherit;margin-top:8px}
#tx #gmodal select{padding:6px 8px;border:1px solid #d1d5db;border-radius:6px}
#tx .tabs{display:flex;gap:4px;margin-bottom:8px}
#tx .find .card.help{width:min(760px,96vw);max-height:84vh;font-size:14px;line-height:1.5}
#tx .help .hhead{display:flex;justify-content:space-between;align-items:center;position:sticky;top:-16px;background:#fff;padding:8px 0;margin-top:-8px;z-index:1}
#tx .help .hidx{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:4px}
#tx .help .hidx a{font-size:12px;padding:3px 10px;border:1px solid #d1d5db;border-radius:14px;text-decoration:none;color:#374151}
#tx .help .hidx a:hover{border-color:var(--brand);color:var(--brand)}
#tx .help h4{margin:18px 0 6px;font-size:15px;color:var(--brand)}
#tx .help p,#tx .help ul,#tx .help dl{margin:0 0 8px}
#tx .help ul{padding-left:20px}
#tx .help li{margin-bottom:4px}
#tx .help dl{display:grid;grid-template-columns:max-content 1fr;gap:6px 12px;align-items:baseline}
#tx .help dt{font-weight:600}
#tx .help dd{margin:0}
#tx .tabs a{padding:6px 14px;border-radius:6px 6px 0 0;font-size:14px;color:#374151;text-decoration:none;border-bottom:2px solid transparent}
#tx .tabs a.on{color:var(--brand);font-weight:600;border-bottom-color:var(--brand)}
#tx #pp{display:none;border:1px solid var(--line);border-radius:10px;background:#fff;padding:12px 14px;min-height:520px}
#tx #pp .tools{padding:0 0 10px;border:none}
#tx #pp .intro{font-size:13px;color:var(--muted);margin:0 0 10px}
#tx .chips{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
#tx .chips a{padding:4px 10px;border:1px solid #d1d5db;border-radius:14px;font-size:12px;color:#374151;text-decoration:none}
#tx .chips a.on{background:var(--brand);border-color:var(--brand);color:#fff}
#tx .chips a.bad{border-color:#f59e0b}
#tx .chips a.bad.on{background:#b45309;border-color:#b45309}
#tx .khelp{font-size:13px;margin:0 0 8px;padding:8px 10px;border-radius:6px;background:#fffbeb;color:#92400e}
#tx table.pt{width:100%;border-collapse:collapse;font-size:13px}
#tx table.pt th{text-align:left;font-size:12px;color:var(--muted);font-weight:600;padding:6px 8px;border-bottom:1px solid var(--line)}
#tx table.pt td{padding:6px 8px;border-bottom:1px solid #f3f4f6;vertical-align:top}
#tx table.pt td a{margin-right:8px}
#tx .p-active{background:#dcfce7;color:#166534}
#tx .p-sub{background:#dcfce7;color:#166534}
#tx .p-none{background:#f3f4f6;color:#6b7280}
@media (max-width:820px){
  #tx .wrap{height:calc(100dvh - 150px);min-height:420px}
  #tx textarea,#tx input,#tx select{font-size:16px}
  #tx .qr{flex-wrap:nowrap;overflow-x:auto;-webkit-overflow-scrolling:touch;padding-bottom:2px}
  #tx .qr button{flex:none}
  #txapp:has(.wrap.open) .apphead{display:none}
  #tx .wrap.open{height:calc(100dvh - 80px)}
  #tx .head{padding:8px 10px;position:relative}
  #tx .head .acts{float:none;position:static;display:block;margin:0}
  #tx .head .acts .back{position:absolute;left:8px;top:8px;padding:4px 10px}
  #tx .head .acts #done,#tx .head .acts #undone{position:absolute;right:8px;top:8px;padding:4px 10px}
  #tx .head h3,#tx .head .sub{margin:0 64px;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  #tx .head h3{font-size:16px}
  #tx .head .sub{font-size:11px}
  #tx .head .ppl{display:none;margin-top:6px}
  #tx .head.showppl .ppl{display:block}
  #tx .pplbtn{display:block;text-align:center;margin-top:2px}
  #tx .qr button{padding:2px 10px;min-height:28px}
  #tx .compose .row{margin-top:2px}
  #tx .xtools{margin-top:4px}
  #tx .compose{padding:8px 10px}
  #tx .help dl{grid-template-columns:1fr;gap:2px}
  #tx .help dd{margin-bottom:6px}
  #tx .m{max-width:86%}
  #tx button{min-height:34px}
  #tx table.pt .wide{display:none}
  #tx .left{width:100%}
  #tx .wrap.open .left{display:none}
  #tx .wrap:not(.open) .right{display:none}
  #tx .back{display:inline-block}
}
</style>"""
        body = """
<div id="tx" style="position:relative">
  <div class="tabs" id="tabs"><a href="#" data-t="inbox" class="on">Conversations</a><a href="#" data-t="people">People</a></div>
  <div id="pp">
    <p class="intro">Who TouchPoint lists as subscribed to the group, next to who PoweredByText actually sends group texts to, matched by cell number.
    This only reads: nothing is sent or changed. A STOP or START texted by the person updates both; adding or removing someone by hand in PoweredByText changes only PoweredByText.</p>
    <div class="tools"><select id="pgrp"></select> <input id="pq" placeholder="Search name or number"> <button class="ghost" id="preload">Refresh</button> <span id="pst" class="tm"></span></div>
    <div class="chips" id="pchips"></div>
    <div id="phelp"></div>
    <div id="ptab"></div>
  </div>
  <div class="wrap" id="wrap">
    <div class="left">
      <div class="tools">
        <div class="seg" id="seg"><a href="#" data-f="new" class="on">New</a><a href="#" data-f="notanswered">Not answered</a><a href="#" data-f="all">All</a><a href="#" data-f="done">Done</a></div>
        <button class="ghost" id="markall" title="Mark every New conversation in this list as read for you">Mark all read</button>
        <button class="ghost" id="newbtn" title="Text someone"__NEWHIDE__>New text</button>
        <button class="ghost" id="gbtn" title="Text everyone in a group"__NEWHIDE__>Group text</button>
        <select id="grp"><option value="">All my groups</option></select>
        <select id="days"><option value="30">30 days</option><option value="90" selected>90 days</option><option value="365">1 year</option></select>
        <input id="q" placeholder="Search name, number or text">
        <a href="#" id="howlink" style="font-size:12px;align-self:center">Help</a>
      </div>
      <div class="list" id="list"><div class="empty">Loading...</div></div>
    </div>
    <div class="right" id="right"><div class="empty">Pick a conversation.</div></div>
  </div>
  <div class="find" id="find"><div class="card">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
      <b>Text someone</b><a href="#" id="findclose">Close</a></div>
    <input id="findq" placeholder="Name or cell phone">
    <div style="font-size:12px;color:#6b7280;margin:6px 0">Only people subscribed to a group you can send from, with texts turned on and not opted out, are listed.</div>
    <div id="findres"></div>
  </div></div>
  <div class="find" id="howmodal"><div class="card help">
    <div class="hhead"><b>Texting help</b><a href="#" id="howclose">Close</a></div>
    <div class="hidx">
      <a href="#" data-go="h-list">The list</a><a href="#" data-go="h-reply">Replying</a><a href="#" data-go="h-new">New text</a>
      <a href="#" data-go="h-group">Group text</a><a href="#" data-go="h-pics">Pictures and files</a><a href="#" data-go="h-links">Links</a>
      <a href="#" data-go="h-cost">What a text costs</a><a href="#" data-go="h-refresh">Updates</a><a href="#" data-go="h-who">Who can see and send</a>
      <a href="#" data-go="h-phone">On your phone</a><a href="#" data-go="h-trouble">When a text will not send</a>
    </div>

    <h4 id="h-list">The list</h4>
    <dl>
      <dt><span class="pill p-new">New</span></dt><dd>They texted since <i>you</i> last opened the conversation. It is yours alone: opening it clears it for you, not for anyone else. <b>Mark all read</b> clears every New one in the list, which helps after a group text that asked for replies.</dd>
      <dt><span class="pill p-notanswered">Not answered</span></dt><dd>They texted last. Everyone sees the same thing.</dd>
      <dt><span class="pill p-answered">Answered</span></dt><dd>Someone on staff texted after their last text, from here, from TouchPoint or from PoweredByText.</dd>
      <dt><span class="pill p-done">Done</span></dt><dd>Anyone can mark a conversation handled. It comes back if they text again.</dd>
      <dt><span class="pill p-optedout">Opted out</span></dt><dd>They texted STOP. Nobody can text them from that group until they text START.</dd>
    </dl>
    <p>Reactions (&ldquo;Liked&hellip;&rdquo;, &#128077;), STOP, START and HELP, group texts and automatic replies never make a conversation New and never count as an answer.</p>
    <p><b>Search</b> finds any part of a first name, goes-by name or last name, in any order: &ldquo;sar whit&rdquo; finds Sarah Whitaker. It also finds a phone number or words in the text. Use the group and day lists to narrow it down.</p>

    <h4 id="h-reply">Replying</h4>
    <ul>
      <li>The reply goes to the person TouchPoint matched the number to, marked &#10003;. When several people share a phone, they all get it on that phone.</li>
      <li>A reply always goes <b>from the group they texted</b>, so it lands in the same conversation on their phone. The line under the box shows the group and any others they are in.</li>
      <li><b>Quick replies</b> fill the box; nothing goes out until you press <b>Send</b>. On a computer, Ctrl+Enter (Cmd+Enter on a Mac) sends too.</li>
      <li><b>Quiet hours:</b> nothing goes out from 9 p.m. to 8 a.m., unless they texted in within the last 2 hours.</li>
    </ul>

    <h4 id="h-new">Starting a new text</h4>
    <p>Press <b>New text</b>, type part of a name or a cell number, then press the group to send from. Only people subscribed to one of your groups, with texts turned on and not opted out, are listed. If they are in more than one group, you can still change the group before the first text goes out.</p>

    <h4 id="h-group">Group text</h4>
    <p>Sends one text to everyone subscribed to the group in PoweredByText, including people who joined by texting a keyword. The window shows how many people will get it and about how many credits it will use, and asks you to confirm. Group texts never go out during quiet hours.</p>

    <h4 id="h-pics">Pictures and files</h4>
    <ul>
      <li>Press <b>Attach</b>. Your administrator decides what is allowed: pictures, PDFs, or small video and audio.</li>
      <li>A picture message is <b>3 credits per person</b>, whatever is in it. Emoji and links add nothing to that.</li>
      <li>Phones can take at most 950 KB. Photos are made smaller for you before they are sent. A file that is still too big is refused; put it on the website and send the link instead.</li>
      <li>Some phones cannot show pictures. They get a plain text instead, with a link to the picture. Press <b>Edit</b> on the line under the box to change that text or leave the link out. A PDF, video or audio file needs the link, or they have no way to open it.</li>
    </ul>

    <h4 id="h-links">Links</h4>
    <p>Paste the full link. Long links are sent as a short PoweredByText link, which saves characters and credits. Do not use bit.ly, Linktree or other public short links: carriers block texts that contain them, and PoweredByText can suspend the account. The page will not send one. RSVP and sign-up links that TouchPoint adds are left alone.</p>

    <h4 id="h-cost">What a text costs</h4>
    <ul>
      <li>Plain text: 1 credit for up to 160 characters, then 1 credit for every 153.</li>
      <li>One emoji, curly quote or some accented letters change the whole text to a mode that fits only 70 characters per credit (67 when longer). The counter under the box shows this, and <b>Make it plain</b> swaps curly quotes and dashes for plain ones.</li>
      <li>Anything over 3 credits goes as a picture message, which is 3 credits.</li>
      <li>A group text multiplies all of this by the number of people, so the window shows the total before you send.</li>
    </ul>

    <h4 id="h-refresh">Updates</h4>
    <p>The page checks for new texts on its own: every 10 seconds while you are working in a conversation, every 15 while you are on the page, and once a minute after 5 minutes with no clicks or typing. It checks right away when you come back. It stops while the browser tab is hidden. Your half-typed reply and any attachment are kept.</p>

    <h4 id="h-who">Who can see and send</h4>
    <p>You see conversations for the texting groups you are listed in on TouchPoint&rsquo;s SMS settings. To send, you also need the SendSMS role and to be a member of the group. Ask your administrator if a group is missing. If your administrator lets admins see every group, admins can read the others but send only from their own.</p>

    <h4 id="h-phone">On your phone</h4>
    <p>This page works in a phone browser. <a href="https://tpxigo.com" target="_blank" rel="noopener">TPxi Go</a>, from the same people who made this page,
    is an app for TouchPoint that brings texting to your phone, with an alert when someone texts in so you can answer right away.</p>

    <h4 id="h-trouble">When a text will not send</h4>
    <dl>
      <dt>Opted out</dt><dd>They texted STOP to this group. Only they can undo it, by texting START.</dd>
      <dt>Texts turned off</dt><dd>Their TouchPoint record says not to text them.</dd>
      <dt>Not subscribed to this group</dt><dd>A new conversation can only start with people subscribed to the group. Pick another group, or ask them to text the group first.</dd>
      <dt>Quiet hours</dt><dd>Wait until 8 a.m., or reply soon after they text in.</dd>
      <dt>Public short link</dt><dd>Replace the bit.ly (or similar) link with the full link.</dd>
      <dt>Emoji are turned off</dt><dd>Your administrator turned emoji off. Take them out.</dd>
      <dt>File too big</dt><dd>Use a smaller picture, or send a link.</dd>
    </dl>
  </div></div>
  <div class="find" id="gmodal"><div class="card">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
      <b>Group text</b><a href="#" id="gclose">Close</a></div>
    <div><select id="gsel"></select> <span id="gcount" class="tm"></span></div>
    <textarea id="gmsg" placeholder="Text everyone in this group"></textarea>
    <div id="gxtools"></div>
    <div id="gmeter" class="meter"></div>
    <div id="gxpanels"></div>
    <div id="gquiet" class="note" style="display:none;margin-top:8px">It is quiet hours (9 p.m. to 8 a.m.). Group texts can't go out until 8 a.m.</div>
    <div style="display:flex;justify-content:space-between;align-items:center;margin-top:10px">
      <span id="gst" style="font-size:12px;color:#6b7280">Goes to everyone subscribed to the group in PoweredByText, including people who joined by keyword.</span>
      <button id="gsend" disabled>Send to group</button></div>
  </div></div>
</div>"""
        js = """
<script>
(function(){
  var ME_FIRST = __FIRST__;
  var POLICY = __POLICY__;     // emoji: allow | replies | none; attach: off | pictures | pictures_pdf | all
  var base = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
  var shownRows = [];
  var threads = [], groups = [], filt = 'new', cur = null, curData = null, busy = false;
  var curGT = {key: null, list: []};   // group texts sent in PoweredByText's app, read once per open
  // Quick replies: plain text (no curly quotes), each one credit. The emoji
  // one switches to the 70-character mode but is still a single credit.
  var QUICK = ['\\ud83d\\udc4d', 'Thanks!', 'Got it, thank you!', 'Thanks, we have you down!', 'See you there!'];
  function $(id){ return document.getElementById(id); }
  function esc(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
  function fmtNum(n){ return n && n.length == 10 ? '(' + n.substr(0,3) + ') ' + n.substr(3,3) + '-' + n.substr(6) : (n || ''); }
  function pad(n){ return (n < 10 ? '0' : '') + n; }
  function nowStr(){ var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate()) + ' ' +
    pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds()); }
  function toDate(t){ return new Date(t.replace(' ', 'T')); }
  function when(t){
    if (!t) return '';
    var d = toDate(t), n = new Date();
    if (d.toDateString() == n.toDateString()) return d.toLocaleTimeString([], {hour:'numeric', minute:'2-digit'});
    var y = new Date(n); y.setDate(n.getDate() - 1);
    if (d.toDateString() == y.toDateString()) return 'Yesterday';
    if ((n - d) < 6 * 864e5) return d.toLocaleDateString([], {weekday:'short'});
    return d.toLocaleDateString([], {month:'short', day:'numeric'});
  }
  function post(data, cb){
    data.now = nowStr();
    var x = new XMLHttpRequest();
    x.open('POST', base, true);
    x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    x.onload = function(){ var r; try { r = JSON.parse(x.responseText); } catch(e) {
      r = {success:false, message:'Unexpected reply from TouchPoint.'}; } cb(r); };
    x.onerror = function(){ cb({success:false, message:'Network error'}); };
    x.send(Object.keys(data).map(function(k){ return k + '=' + encodeURIComponent(data[k]); }).join('&'));
  }
  // ------------------------------------------------ credit meter
  // Plain text (GSM-7): 160 characters for one credit, then 153 per credit.
  // Any other character (emoji, curly quotes) makes it Unicode: 70, then 67.
  var GSM = '@\\u00a3$\\u00a5\\u00e8\\u00e9\\u00f9\\u00ec\\u00f2\\u00c7\\n\\u00d8\\u00f8\\r\\u00c5\\u00e5\\u0394_\\u03a6\\u0393\\u039b\\u03a9\\u03a0\\u03a8\\u03a3\\u0398\\u039e\\u00c6\\u00e6\\u00df\\u00c9 !"#\\u00a4%&\\'()*+,-./0123456789:;<=>?\\u00a1ABCDEFGHIJKLMNOPQRSTUVWXYZ\\u00c4\\u00d6\\u00d1\\u00dc\\u00a7\\u00bfabcdefghijklmnopqrstuvwxyz\\u00e4\\u00f6\\u00f1\\u00fc\\u00e0';
  var GSMX = '^{}\\\\[~]|\\u20ac';
  var PLAIN = {'\\u2018':"'", '\\u2019':"'", '\\u201a':"'", '\\u201b':"'", '\\u2032':"'", '\\u201c':'"', '\\u201d':'"', '\\u201e':'"',
               '\\u2033':'"', '\\u2013':'-', '\\u2014':'-', '\\u2012':'-', '\\u2010':'-', '\\u2011':'-', '\\u2026':'...', '\\u00a0':' ',
               '\\u2002':' ', '\\u2003':' ', '\\u2009':' ', '\\u200b':'', '\\u2022':'-', '\\u00b7':'-'};
  function makePlain(t){ return Array.from(t).map(function(c){ return PLAIN.hasOwnProperty(c) ? PLAIN[c] : c; }).join(''); }
  function credits(t){
    var sept = 0, bad = [];
    Array.from(t).forEach(function(c){
      if (GSM.indexOf(c) >= 0) sept += 1;
      else if (GSMX.indexOf(c) >= 0) sept += 2;
      else if (bad.indexOf(c) < 0) bad.push(c);
    });
    var uni = bad.length > 0, n = uni ? t.length : sept;   // t.length counts an emoji as 2
    var segs = !n ? 0 : uni ? (n <= 70 ? 1 : Math.ceil(n / 67)) : (n <= 160 ? 1 : Math.ceil(n / 153));
    return {n: n, chars: Array.from(t).length, segs: segs, uni: uni, bad: bad};
  }
  // Renders the meter into el. people = recipients for a group text.
  // onPlain(newBody) is called when "Make it plain" is pressed.
  function renderMeter(el, full, body, people, onPlain){
    if (!full) { el.innerHTML = ''; return null; }
    var c = credits(full), h = [];
    var s1 = c.chars + ' characters &middot; <b>' + c.segs + ' credit' + (c.segs == 1 ? '' : 's') + '</b>' + (people != null ? ' each' : '');
    if (people != null) s1 += ' &times; ' + people + ' people = <b>about ' + (c.segs * people).toLocaleString() + ' credits</b>';
    h.push('<div>' + s1 + '</div>');
    if (c.uni) {
      var fixed = makePlain(full), fc = credits(fixed), canFix = fixed != full;
      var shown = c.bad.slice(0, 10).map(function(ch){ return '<span class="ch">' + esc(ch == '\\u00a0' ? ' ' : ch) + '</span>'; }).join('');
      var line = 'These characters switch the text to the costly mode (70 characters per credit instead of 160): ' + shown + '.';
      if (canFix && fc.segs < c.segs) {
        var save = c.segs - fc.segs;
        line += ' Plain quotes and dashes would make it ' + fc.segs + ' credit' + (fc.segs == 1 ? '' : 's') +
          (people != null ? ', saving about ' + (save * people).toLocaleString() + ' credits' : '') + '.';
      }
      if (canFix) line += '<button type="button" class="ghost mkplain">Make it plain</button>';
      h.push('<div class="warn">' + line + '</div>');
    } else if (c.segs > 1) {
      var left = c.segs * 153 - c.n;
      h.push('<div>Long text: ' + c.segs + ' credits each. ' + (c.n - (c.segs - 1) * 153) + ' characters into the last credit' +
        (c.segs > 3 ? '. A picture message is 3 credits flat' : '') + '.</div>');
    }
    el.innerHTML = h.join('');
    var btn = el.querySelector('.mkplain');
    if (btn) btn.onclick = function(){ onPlain(makePlain(body)); };
    return c;
  }

  function store(k, v){ try { if (v === undefined) return localStorage.getItem('txinbox.' + k); localStorage.setItem('txinbox.' + k, v); } catch(e) { return null; } }

  // ------------------------------------------------ list
  function loadThreads(){
    post({action:'threads', days:$('days').value}, function(r){
      if (!r.success) { $('list').innerHTML = '<div class="empty">' + esc(r.message) + '</div>'; return; }
      threads = r.threads;
      if (!groups.length) {
        groups = r.groups;
        $('grp').innerHTML = '<option value="">All my groups</option>' + groups.map(function(g){
          return '<option value="' + g.id + '">' + esc(g.name) + '</option>'; }).join('');
        var saved = store('grp'); if (saved) $('grp').value = saved;
        // Which view shows (Conversations or Subscribers) is decided by the page frame, not remembered here.
        // Reopen the conversation this tab had open before a refresh.
        var open = tabKey();
        if (open && !cur) { var k = open.split('|'); openThread(parseInt(k[0], 10), k[1], 'read'); }
      }
      renderList();
    });
  }
  // The open conversation is remembered for this browser tab only (session
  // storage), so a refresh comes back to it. Nothing goes in the address,
  // so no phone number ends up in history or a copied link.
  function tabKey(v){
    try { if (v === undefined) return sessionStorage.getItem('txinbox.open') || ''; if (v) sessionStorage.setItem('txinbox.open', v); else sessionStorage.removeItem('txinbox.open'); }
    catch(e) { return ''; }
  }
  var LBL = {notanswered:'Not answered', done:'Done', optedout:'Opted out', answered:'Answered', other:''};
  function renderList(){
    var g = $('grp').value, q = ($('q').value || '').toLowerCase();
    var rows = threads.filter(function(t){
      if (g && String(t.gid) != g) return false;
      if (filt == 'new' && !t.isNew) return false;
      if (filt == 'notanswered' && t.status != 'notanswered') return false;
      if (filt == 'done' && t.status != 'done') return false;
      if (q) { var hay = (t.name + ' ' + t.num + ' ' + t.preview + ' ' + t.group).toLowerCase();
        if (!q.split(/\\s+/).every(function(w){ return !w || hay.indexOf(w) >= 0; })) return false; }
      return true;
    });
    var n = threads.filter(function(t){ return t.isNew && (!g || String(t.gid) == g); }).length;
    $('seg').firstChild.textContent = 'New' + (n ? ' (' + n + ')' : '');
    $('markall').disabled = !rows.some(function(t){ return t.isNew; });
    shownRows = rows;
    if (!rows.length) {
      $('list').innerHTML = '<div class="empty">' + (filt == 'new' ? 'Nothing new for you.<br><span class="tm">Older conversations are under All.</span>'
        : threads.length ? 'No conversations match.' : 'No conversations yet.<br><span class="tm">When someone texts one of your groups, it shows up here.</span>') + '</div>';
      return;
    }
    $('list').innerHTML = rows.map(function(t){
      return '<div class="th' + (t.isNew ? ' isnew' : '') + (cur && cur.key == t.key ? ' on' : '') + '" data-k="' + esc(t.key) + '">' +
        '<div class="top"><span class="nm">' + esc(t.name || fmtNum(t.num)) + (t.others > 0 ? ' <span class="tm">+' + t.others + '</span>' : '') +
        '</span><span class="tm">' + when(t.t) + '</span></div>' +
        '<div class="pv">' + (t.lastDir == 'out' ? 'Staff: ' : t.lastDir == 'bc' ? 'Group text: ' : t.lastDir == 'auto' ? 'Automatic reply: ' : '') + esc(t.preview) + '</div>' +
        '<div class="meta"><span>' + esc(t.group) + '</span>' +
        (t.isNew ? '<span class="pill p-new">New' + (t.newCount > 1 ? ' ' + t.newCount : '') + '</span>' : '') +
        (LBL[t.status] ? '<span class="pill p-' + t.status + '">' + LBL[t.status] + '</span>' : '') +
        '</div></div>';
    }).join('');
    Array.prototype.forEach.call(document.querySelectorAll('#list .th'), function(el){
      el.onclick = function(){
        var t = threads.filter(function(x){ return x.key == el.getAttribute('data-k'); })[0];
        if (t) openThread(t.gid, t.num, true);
      };
    });
  }

  // ------------------------------------------------ thread
  function openThread(gid, num, claim){
    cur = {key: gid + '|' + num, gid: gid, num: num};
    tabKey(cur.key);
    $('wrap').className = 'wrap open';
    if (claim === true) { $('right').innerHTML = '<div class="empty">Loading...</div>'; renderList(); }
    var key = gid + '|' + num, wantGT = curGT.key != key;
    post({action:'thread', gid:gid, num:num, claim:claim === 'read' ? 'read' : claim ? '1' : '0', gt:wantGT ? '1' : '0'}, function(r){
      if (!cur || cur.key != key) return;
      if (!r.success) { $('right').innerHTML = '<div class="empty">' + esc(r.message) + '</div>'; return; }
      if (r.groupTexts) curGT = {key: key, list: r.groupTexts};
      var keep = $('msg') ? $('msg').value : '';
      var atBottom = !$('msgs') || ($('msgs').scrollHeight - $('msgs').scrollTop - $('msgs').clientHeight < 40);
      curData = r; renderThread(keep, claim === true || atBottom);
    });
  }
  function renderThread(draft, scroll){
    var r = curData, html = [];
    var ppl = r.people.map(function(p){
      return '<a href="/Person2/' + p.pid + '" target="_blank"' + (p.adult ? '' : ' class="kid"') + '>' + esc(p.name) +
        (p.age != null ? ' (' + p.age + ')' : '') + '</a>' + (p.target ? '&#10003; ' : '') +
        (!p.receive ? '<span class="pill p-optedout">texts off</span> ' : '') + (p.opted ? '<span class="pill p-optedout">opted out</span> ' : '');
    }).join('');
    html.push('<div class="head"><div class="acts"><button class="ghost back" id="back">Back</button>' +
      (r.done ? '<button class="ghost" id="undone">Not done</button>' : '<button class="ghost" id="done">Done</button>') + '</div>' +
      '<h3>' + esc(r.who || fmtNum(r.num)) + '</h3><div class="sub">' + fmtNum(r.num) + ' &middot; ' + esc(r.group) +
      (r.done ? ' &middot; Done by ' + esc(r.doneBy) + ' ' + when(r.done) : '') +
      '</div>' +
      (r.people.length > 1 ? '<a href="#" class="pplbtn" id="pplbtn">' + r.people.length + ' people on this number</a>' : '') +
      '<div class="ppl">' + (ppl || '<span class="tm">No one in TouchPoint has this number.</span>') + '</div></div>');
    var evs = r.events.slice();
    if (curGT.key == cur.key) curGT.list.forEach(function(g){
      if (!evs.some(function(e){ return e.dir == 'out' && e.body == g.body && Math.abs(toDate(e.t) - toDate(g.t)) < 36e5; }))
        evs.push({dir:'out', t:g.t, body:g.body, broadcast:true, app:true, status:''});
    });
    evs.sort(function(a, b){ return a.t < b.t ? -1 : a.t > b.t ? 1 : 0; });
    var lastDay = '', msgs = evs.map(function(e){
      var d = toDate(e.t).toDateString(), sep = '';
      if (d != lastDay) { lastDay = d; sep = '<div class="day">' + toDate(e.t).toLocaleDateString([], {weekday:'long', month:'long', day:'numeric'}) + '</div>'; }
      var cls = e.dir + (e.broadcast ? ' bc' : '') + (e.auto ? ' auto' : '') + (e.reaction ? ' rx' : '') + (e.stop ? ' stop' : '');
      var info = toDate(e.t).toLocaleTimeString([], {hour:'numeric', minute:'2-digit'});
      if (e.dir == 'out') info = (e.auto ? 'Automatic reply' : e.app ? 'Group text sent in PoweredByText' : e.broadcast ? 'Group text' + (e.staff && e.staff != 'PoweredByText' ? ' from ' + esc(e.staff) : '')
        : esc(e.staff || '')) + ' &middot; ' + info +
        (e.status && !/deliver/i.test(e.status) ? ' &middot; ' + esc(e.status) : '');
      return sep + '<div class="m ' + cls + '"><div class="b">' + esc(e.body) + '</div><div class="i">' + info + '</div></div>';
    }).join('');
    html.push('<div class="msgs" id="msgs">' + (msgs || '<div class="empty">No messages yet.</div>') + '</div>');
    if (!r.canSend) {
      html.push('<div class="note red">' + esc(r.reason) + '</div>');
    } else {
      if (r.quiet) html.push('<div class="note">' + esc(r.quiet) + '</div>');
      if (r.via == 'pbt') html.push('<div class="note">No one in TouchPoint has this number, so this goes out through PoweredByText directly.</div>');
      html.push('<div class="compose"><div class="qr" id="qr">' + QUICK.map(function(q, i){
        return '<button type="button" data-i="' + i + '" title="Put this in the box. Nothing is sent until you press Send.">' + esc(q) + '</button>'; }).join('') + '</div>' +
        '<div class="crow"><textarea id="msg" rows="1" placeholder="Text ' + esc(r.who || fmtNum(r.num)) + '"></textarea>' +
        '<button id="send"' + (r.quiet ? ' disabled' : '') + '>Send</button></div>' + curX().toolbar() +
        '<div class="row"><span id="sendst"></span></div>' +
        '<div class="meter" id="count"></div>' + curX().panels() + '</div>');
    }
    $('right').innerHTML = html.join('');
    if (scroll && $('msgs')) $('msgs').scrollTop = $('msgs').scrollHeight;
    $('back').onclick = function(){ cur = null; tabKey(''); $('wrap').className = 'wrap'; renderList(); };
    if ($('done')) $('done').onclick = function(){ markDone(false); };
    if ($('undone')) $('undone').onclick = function(){ markDone(true); };
    if ($('msg')) {
      $('msg').value = draft || '';
      curX().bind($('msg'));
      drawFrom(r);
      $('msg').oninput = function(){ grow(); count(); }; count(); grow();
      if ($('pplbtn')) $('pplbtn').onclick = function(e){ e.preventDefault(); this.parentNode.classList.toggle('showppl'); };
      $('send').onclick = send;
      Array.prototype.forEach.call(document.querySelectorAll('#qr button'), function(b){
        b.onclick = function(){ $('msg').value = QUICK[+b.getAttribute('data-i')]; count(); $('msg').focus(); };
      });
      $('msg').onkeydown = function(e){ if (e.key == 'Enter' && (e.metaKey || e.ctrlKey)) send(); };
    }
  }
  // The reply box starts one line tall and grows with what is typed.
  function grow(){ var t = $('msg'); if (!t) return; t.style.height = 'auto'; t.style.height = Math.min(t.scrollHeight + 2, 160) + 'px'; }
  function fullText(){
    var t = ($('msg').value || '').trim();
    return t;
  }
  function count(){
    var X = curX(), full = fullText();
    var ok = X.draw($('msg').value, false);
    if (!X.meter($('count'), full, null))
      renderMeter($('count'), asSent(full), $('msg').value, null, function(plain){ $('msg').value = plain; count(); });
    if ($('send')) $('send').disabled = !ok || !!(curData && curData.quiet);
  }
  // Replies go from the group they texted, so the answer lands in the same
  // conversation on their phone. A new conversation can switch groups.
  function drawFrom(r){
    var el = $('txfrom'), also = r.alsoIn || [];
    if (!el) return;
    if (!r.events.length && also.length) {
      el.innerHTML = 'From <select id="txfromsel">' + [{gid: r.gid, name: r.group}].concat(also).map(function(g){
        return '<option value="' + g.gid + '"' + (g.gid == r.gid ? ' selected' : '') + '>' + esc(g.name) + '</option>'; }).join('') + '</select>';
      $('txfromsel').onchange = function(){ openThread(parseInt(this.value, 10), r.num, true); };
    } else {
      el.innerHTML = 'From <b>' + esc(r.group) + '</b>' + (also.length ? ' <span title="Replies go from the group they texted, so the answer lands in the same conversation on their phone.">(also in '
        + also.map(function(g){ return esc(g.name); }).join(', ') + ')</span>' : '');
    }
  }
  function send(){
    if (busy || !cur) return;
    var t = fullText(), X = curX();
    if (!t && !X.st.media) return;
    if (!X.draw($('msg').value, false)) return;
    busy = true; $('send').disabled = true; $('sendst').textContent = 'Sending...';
    var data = {action:'send', gid:cur.gid, num:cur.num, msg:t}, pl = X.payload(t);
    Object.keys(pl).forEach(function(k){ data[k] = pl[k]; });
    post(data, function(r){
      busy = false;
      if (!r.success) { $('send').disabled = false; $('sendst').textContent = r.message; return; }
      $('msg').value = ''; X.clear(); $('sendst').textContent = 'Sent' + (r.note ? '. ' + r.note : '');
      setTimeout(function(){ if (cur) openThread(cur.gid, cur.num, false); loadThreads(); }, 1200);
    });
  }
  function markDone(undo){
    post({action:'done', gid:cur.gid, num:cur.num, undo:undo ? '1' : '0'}, function(r){
      if (!r.success) { alert(r.message); return; }
      loadThreads(); openThread(cur.gid, cur.num, false);
    });
  }

  // ------------------------------------------------ pictures, emoji and links
  // One set of these per place you can type: each conversation keeps its own
  // (so the 15-second refresh does not lose an attachment) and the group text
  // window has one.
  var EMOJI = ['👍', '🙏', '❤️', '😊', '🎉', '✅', '📅', '📍',
               '⛪', '🙌', '😀', '👋', '⏰', '📣', '☀️', '🍕',
               '🎶', '📖', '🚌', '🎈'];
  var MEDIA_MAX = 950 * 1024, SHRINK_TO = 900 * 1024;
  var ACCEPT = {off: '', pictures: 'image/jpeg,image/png,image/gif',
                pictures_pdf: 'image/jpeg,image/png,image/gif,application/pdf',
                all: 'image/jpeg,image/png,image/gif,application/pdf,video/mp4,audio/mpeg'};
  var PUBLIC_RE = /https?:\\/\\/(www\\.)?(bit\\.ly|tinyurl\\.com|linktr\\.ee|linkr\\.ee|t\\.co|ow\\.ly|is\\.gd|rebrand\\.ly|shorturl\\.at|cutt\\.ly|tiny\\.cc|buff\\.ly|t\\.ly|rb\\.gy)\\//i;
  var ALREADY_RE = /https?:\\/\\/([\\w-]+\\.)*(gopbt\\.me|tpsdb\\.co)\\//i;
  var URL_RE = /https?:\\/\\/[\\w-]+\\.[^\\s<>"]+/g;
  var EMOJI_RE = /[\\ud800-\\udbff][\\udc00-\\udfff]|[☀-➿⬀-⯿️‍]/;
  function longLink(u){ return !PUBLIC_RE.test(u) && !ALREADY_RE.test(u) && u.length > 32; }
  // The text as it will go out: long links become a PBT short link (about 26 characters).
  function asSent(t){ return (t || '').replace(URL_RE, function(u){ return longLink(u) ? 'https://q.gopbt.me/abc123' : u; }); }
  function mediaWord(mime){ return /^image/.test(mime) ? 'Picture' : mime == 'application/pdf' ? 'File' : 'Media'; }
  function Extras(p, gidFn, onChange){
    var st = {media: null, mime: '', name: '', kb: 0, note: '', err: '', fb: '', fbEdited: false, keep: true, open: false, busy: false};
    var self = {st: st};
    function $p(id){ return $(p + id); }
    self.toolbar = function(){
      var acc = ACCEPT[POLICY.attach] || '';
      return '<div class="xtools">'
        + (acc ? '<button type="button" class="ghost xbtn" id="' + p + 'att" title="Picture message: 3 credits, emoji and links included">Attach</button>'
               + '<input type="file" id="' + p + 'file" accept="' + acc + '" style="display:none">' : '')
        + (POLICY.emoji != 'none' ? '<button type="button" class="ghost xbtn" id="' + p + 'emo" title="Emoji">🙂</button>' : '')
        + '<span class="xchip" id="' + p + 'chip" style="display:none"></span>'
        + '<span class="xfrom" id="' + p + 'from"></span></div>'
        + '<div class="xpick" id="' + p + 'pick" style="display:none">' + EMOJI.map(function(e, i){
            return '<button type="button" data-i="' + i + '">' + e + '</button>'; }).join('') + '</div>';
    };
    self.panels = function(){
      return '<div class="xline" id="' + p + 'links"></div>'
        + '<div class="xline" id="' + p + 'fbline" style="display:none"></div>'
        + '<div class="xfb" id="' + p + 'fbbox" style="display:none"><textarea id="' + p + 'fbtext" aria-label="Text sent if the attachment cannot go"></textarea>'
        + '<div class="xline"><span id="' + p + 'fbcount"></span> <a href="#" id="' + p + 'droplink"></a> <a href="#" id="' + p + 'fbdone">Done</a></div></div>';
    };
    self.bind = function(ta){
      var att = $p('att'), file = $p('file'), emo = $p('emo');
      if (att) att.onclick = function(){ file.value = ''; file.click(); };
      if (file) file.onchange = function(){ if (file.files[0]) self.attach(file.files[0]); };
      if (emo) emo.onclick = function(){ var k = $p('pick'); k.style.display = k.style.display == 'none' ? 'flex' : 'none'; };
      Array.prototype.forEach.call(document.querySelectorAll('#' + p + 'pick button'), function(b){
        b.onclick = function(){
          var e = EMOJI[+b.getAttribute('data-i')], s0 = ta.selectionStart || ta.value.length, s1 = ta.selectionEnd || s0;
          ta.value = ta.value.slice(0, s0) + e + ta.value.slice(s1);
          ta.selectionStart = ta.selectionEnd = s0 + e.length; ta.focus(); onChange();
        };
      });
      $p('fbtext').value = st.fb;
      $p('fbtext').oninput = function(){ st.fb = this.value; st.fbEdited = true; onChange(); };
      $p('fbdone').onclick = function(e){ e.preventDefault(); st.open = false; onChange(); };
      $p('droplink').onclick = function(e){ e.preventDefault(); st.keep = !st.keep; onChange(); };
      self.draw(ta.value, false);
    };
    self.clear = function(){ st.media = null; st.mime = ''; st.name = ''; st.note = ''; st.fb = ''; st.fbEdited = false; st.keep = true; st.open = false; };
    // Draws the chip, fallback line and link notes. Returns false when the
    // text cannot be sent as it is (emoji turned off, public short link).
    self.draw = function(text, isGroup){
      var ok = true, chip = $p('chip');
      if (!chip) return true;
      if (st.busy) { chip.style.display = ''; chip.innerHTML = 'Preparing ' + esc(st.name) + '...'; }
      else if (st.media) {
        chip.style.display = '';
        chip.innerHTML = esc(mediaWord(st.mime)) + ': ' + esc(st.name) + ' (' + st.kb + ' KB' + (st.note ? ', ' + esc(st.note) : '') + ')'
          + ' <button type="button" class="xx" title="Remove">&times;</button>';
        chip.querySelector('.xx').onclick = function(){ self.clear(); onChange(); };
      } else if (st.err) { chip.style.display = ''; chip.innerHTML = '<span class="xbad">' + esc(st.err) + '</span>'; }
      else chip.style.display = 'none';
      var notes = [];
      (text.match(URL_RE) || []).forEach(function(u){
        if (PUBLIC_RE.test(u)) { ok = false; notes.push('<span class="xbad">' + esc(u) + ' is a public short link. Carriers block these and PoweredByText warns it can get the account suspended. Paste the full link instead and it will be shortened safely.</span>'); }
        else if (longLink(u)) notes.push('Link goes out as a short PoweredByText link (about 26 characters).');
      });
      if (EMOJI_RE.test(text) && !st.media) {
        if (POLICY.emoji == 'none') { ok = false; notes.push('<span class="xbad">Emoji are turned off for texts here. Take them out to send.</span>'); }
        else if (isGroup && POLICY.emoji == 'replies') notes.push('<span class="xwarn">Emoji switch this group text to the costly mode for every person it goes to. Attach a picture and they are free.</span>');
      }
      $p('links').innerHTML = notes.filter(function(x, i){ return notes.indexOf(x) == i; }).map(function(x){ return '<div>' + x + '</div>'; }).join('');
      if (!st.fbEdited) st.fb = text;
      var fbl = $p('fbline'), box = $p('fbbox');
      if (st.media) {
        var c = credits(asSent(self.fallback(text)));
        fbl.style.display = st.open ? 'none' : '';
        fbl.innerHTML = 'If their phone cannot show it, they get ' + (st.fbEdited ? 'the text you wrote' : 'the text') + (st.keep ? ' and a link to the ' + mediaWord(st.mime).toLowerCase() : '')
          + ' (' + c.segs + ' credit' + (c.segs == 1 ? '' : 's') + '). <a href="#" class="xedit">Edit</a>';
        fbl.querySelector('.xedit').onclick = function(e){ e.preventDefault(); st.open = true; onChange(); $p('fbtext').focus(); };
        box.style.display = st.open ? '' : 'none';
        $p('fbcount').textContent = c.chars + ' characters, ' + c.segs + ' credit' + (c.segs == 1 ? '' : 's') + (st.keep ? ', link added at the end' : '');
        $p('droplink').textContent = st.keep ? 'Leave out the link' : 'Add the link back';
        if (!st.keep && st.mime && !/^image/.test(st.mime)) $p('fbcount').textContent += '. Without the link they have no way to open the file.';
      } else { fbl.style.display = 'none'; box.style.display = 'none'; }
      if (st.busy) ok = false;
      return ok;
    };
    self.fallback = function(text){ return (st.fbEdited ? st.fb : text) + (st.keep && st.media ? '\\n' + mediaWord(st.mime) + ': https://q.gopbt.me/abc123' : ''); };
    self.payload = function(text){
      return st.media ? {media: JSON.stringify(st.media), fb: st.fbEdited ? st.fb : text, fblink: st.keep ? '1' : '0'} : {};
    };
    self.meter = function(el, full, people){
      if (!st.media) return null;
      var each = 3, fb = credits(asSent(self.fallback(full)));
      el.innerHTML = '<div>' + mediaWord(st.mime) + ' message: <b>3 credits</b>' + (people != null ? ' each &times; ' + people + ' people = <b>about ' + (3 * people).toLocaleString() + ' credits</b>' : '')
        + '. Emoji and links do not add to it. If it falls back to text: ' + fb.segs + ' credit' + (fb.segs == 1 ? '' : 's') + (people != null ? ' each' : '') + '.</div>';
      return {segs: each};
    };
    self.attach = function(f){
      st.err = ''; st.media = null;
      var mime = (f.type || '').toLowerCase();
      if ((ACCEPT[POLICY.attach] || '').split(',').indexOf(mime) < 0) { st.err = 'That kind of file cannot be sent here.'; onChange(); return; }
      st.name = f.name; st.mime = mime; st.busy = true; onChange();
      function done(blob, note){
        if (blob.size > MEDIA_MAX) {
          st.busy = false;
          st.err = f.name + ' is ' + Math.round(blob.size / 1024) + ' KB. A picture message can be at most 950 KB. '
            + (/^image\\/gif/.test(mime) ? 'Use a still picture, or send a link.' : /^image/.test(mime) ? 'Try a smaller picture.' : 'Put it on the website and send the link.');
          onChange(); return;
        }
        var rd = new FileReader();
        rd.onload = function(){
          var b64 = String(rd.result).split(',')[1];
          post({action:'upload_media', gid:gidFn(), fname:f.name.replace(/\\.(png|jpe?g)$/i, '') + (blob.type == 'image/jpeg' ? '.jpg' : f.name.replace(/^.*(\\.[^.]+)$/, '$1')),
                fmime:blob.type || mime, fdata:b64}, function(r){
            st.busy = false;
            if (!r.success) { st.err = r.message; onChange(); return; }
            st.media = r.media; st.mime = blob.type || mime; st.kb = Math.round(blob.size / 1024); st.note = note || ''; onChange();
          });
        };
        rd.readAsDataURL(blob);
      }
      if ((mime == 'image/jpeg' || mime == 'image/png') && f.size > SHRINK_TO) shrink(f, function(b){ done(b, 'shrunk from ' + Math.round(f.size / 1024) + ' KB'); });
      else done(f, '');
    };
    return self;
  }
  // Phone photos are 2 to 5 MB; MMS allows 950 KB. Redraw at most 1600 px on
  // the long side and step the JPEG quality down until it fits.
  function shrink(file, cb){
    var img = new Image(), url = URL.createObjectURL(file);
    img.onload = function(){
      URL.revokeObjectURL(url);
      var sides = [1600, 1280, 1024], qs = [0.85, 0.75, 0.65, 0.55], si = 0, qi = 0;
      (function step(){
        var k = Math.min(1, sides[si] / Math.max(img.width, img.height)), c = document.createElement('canvas');
        c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
        var g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height);
        c.toBlob(function(b){
          if (b.size <= SHRINK_TO || (si == sides.length - 1 && qi == qs.length - 1)) return cb(b);
          if (qi < qs.length - 1) qi++; else { qi = 0; si++; }
          step();
        }, 'image/jpeg', qs[qi]);
      })();
    };
    img.onerror = function(){ URL.revokeObjectURL(url); cb(file); };
    img.src = url;
  }
  var threadX = {};
  function curX(){ if (!cur) return null; return threadX[cur.key] = threadX[cur.key] || Extras('tx', function(){ return cur.gid; }, function(){ if ($('msg')) count(); }); }
  var groupX = Extras('gx', function(){ return $('gsel').value; }, function(){ gUpdate(); });

  // ------------------------------------------------ new text
  var findTimer = null;
  $('newbtn').onclick = function(){ $('find').style.display = 'flex'; $('findq').value = ''; $('findres').innerHTML = ''; $('findq').focus(); };
  $('findclose').onclick = function(e){ e.preventDefault(); $('find').style.display = 'none'; };
  $('findq').oninput = function(){
    clearTimeout(findTimer);
    var term = this.value;
    findTimer = setTimeout(function(){
      post({action:'find', term:term}, function(r){
        if (!r.success) { $('findres').textContent = r.message; return; }
        $('findres').innerHTML = r.people.map(function(p){
          var opts = !p.receive ? '<span class="pill p-optedout">texts turned off</span>' :
            (p.groups.length ? p.groups.map(function(g){ return '<button data-g="' + g.gid + '" data-n="' + p.num + '">' + esc(g.name) + '</button>'; }).join('')
             : '<span class="tm">Not subscribed to any of your groups</span>');
          return '<div class="pr"><b>' + esc(p.name) + '</b>' + (p.age != null ? ' (' + p.age + ')' : '') + ' &middot; ' + fmtNum(p.num) + '<div>' + opts + '</div></div>';
        }).join('') || '<div class="tm">No one found.</div>';
        Array.prototype.forEach.call(document.querySelectorAll('#findres button'), function(b){
          b.onclick = function(){ $('find').style.display = 'none'; openThread(parseInt(b.getAttribute('data-g'), 10), b.getAttribute('data-n'), true); };
        });
      });
    }, 300);
  };

  // ------------------------------------------------ group text
  var gActive = null;
  function gFull(){
    var t = ($('gmsg').value || '').trim();
    return t;
  }
  function isQuiet(){ var h = new Date().getHours(); return h < 8 || h >= 21; }
  function gUpdate(){
    var ok = !$('gxtools').innerHTML || groupX.draw($('gmsg').value, true);
    var c = groupX.meter($('gmeter'), gFull(), gActive) ||
      renderMeter($('gmeter'), asSent(gFull()), $('gmsg').value, gActive, function(plain){ $('gmsg').value = plain; gUpdate(); });
    $('gquiet').style.display = isQuiet() ? '' : 'none';
    $('gsend').disabled = busy || !ok || !(c || groupX.st.media) || !gActive || isQuiet() || gFull().length > 1600;
  }
  function gLoad(){
    gActive = null; $('gcount').textContent = 'Counting subscribers...'; gUpdate();
    var gid = $('gsel').value;
    post({action:'group_info', gid:gid}, function(r){
      if ($('gsel').value != gid) return;
      if (!r.success) { $('gcount').textContent = r.message; return; }
      gActive = r.active;
      var other = Object.keys(r.counts).filter(function(k){ return k.toLowerCase() != 'active'; })
        .map(function(k){ return r.counts[k] + ' ' + k.toLowerCase(); });
      $('gcount').textContent = r.active + ' will get it' + (other.length ? ' (' + other.join(', ') + ' will not)' : '');
      gUpdate();
    });
  }
  $('gbtn').onclick = function(){
    var mine = groups.filter(function(g){ return g.canSend; });
    if (!mine.length) return;
    $('gsel').innerHTML = mine.map(function(g){ return '<option value="' + g.id + '">' + esc(g.name) + '</option>'; }).join('');
    var pref = $('grp').value; if (pref && mine.some(function(g){ return String(g.id) == pref; })) $('gsel').value = pref;
    $('gst').textContent = 'Goes to everyone subscribed to the group in PoweredByText, including people who joined by keyword.';
    if (!$('gxtools').innerHTML) {
      $('gxtools').innerHTML = groupX.toolbar(); $('gxpanels').innerHTML = groupX.panels(); groupX.bind($('gmsg'));
    }
    $('gmodal').style.display = 'flex'; $('gmsg').focus(); gLoad();
  };
  $('markall').onclick = function(){
    var keys = shownRows.filter(function(t){ return t.isNew; }).map(function(t){ return t.key; });
    if (!keys.length) return;
    if (!confirm('Mark ' + keys.length + ' conversation' + (keys.length == 1 ? '' : 's') + ' as read for you?')) return;
    post({action:'mark_read', keys:keys.join(',')}, function(r){ if (!r.success) alert(r.message); loadThreads(); });
  };
  $('howlink').onclick = function(e){ e.preventDefault(); $('howmodal').style.display = 'flex'; };
  $('howclose').onclick = function(e){ e.preventDefault(); $('howmodal').style.display = 'none'; };
  $('howmodal').onclick = function(e){
    var go = e.target.getAttribute && e.target.getAttribute('data-go');
    if (go) { e.preventDefault(); $(go).scrollIntoView({behavior: 'smooth', block: 'start'}); }
    else if (e.target == this) this.style.display = 'none';
  };
  $('gclose').onclick = function(e){ e.preventDefault(); $('gmodal').style.display = 'none'; };
  $('gsel').onchange = gLoad;
  $('gmsg').oninput = gUpdate;
  $('gsend').onclick = function(){
    var t = gFull(), c = groupX.st.media ? {segs: 3} : credits(asSent(t)), name = $('gsel').options[$('gsel').selectedIndex].text;
    if ((!t && !groupX.st.media) || !gActive || busy) return;
    if (!confirm('Send this to ' + gActive + ' people in ' + name + '?\\n\\n' + c.segs + ' credit' + (c.segs == 1 ? '' : 's') +
        ' each, about ' + (c.segs * gActive).toLocaleString() + ' credits in all.\\n\\nThis cannot be undone.')) return;
    busy = true; $('gsend').disabled = true; $('gst').textContent = 'Sending...';
    var data = {action:'group_send', gid:$('gsel').value, msg:t, expected:gActive}, pl = groupX.payload(t);
    Object.keys(pl).forEach(function(k){ data[k] = pl[k]; });
    post(data, function(r){
      busy = false;
      if (!r.success) { $('gst').textContent = r.message; gUpdate(); return; }
      $('gmsg').value = ''; groupX.clear(); gUpdate(); $('gmodal').style.display = 'none';
      setTimeout(loadThreads, 3000);
    });
  };

  // ------------------------------------------------ people (compare, read only)
  var pData = null, pKind = '', LIMIT = 400;
  var TPL = {sub:'Subscribed', opted:'Opted out', '':'Not listed'};
  function pLoad(){
    var gid = $('pgrp').value;
    if (!gid) return;
    pData = null; $('pst').textContent = 'Reading both lists...'; $('pchips').innerHTML = ''; $('phelp').innerHTML = ''; $('ptab').innerHTML = '';
    post({action:'people', gid:gid}, function(r){
      if ($('pgrp').value != gid) return;
      if (!r.success) { $('pst').textContent = r.message; return; }
      pData = r; $('pst').textContent = r.rows.length.toLocaleString() + ' numbers';
      var bad = r.kinds.filter(function(k){ return !/^ok/.test(k.key) && r.counts[k.key]; });
      if (!pKind || !r.counts[pKind]) pKind = bad.length ? bad[0].key : 'okon';
      pRender();
    });
  }
  function pRender(){
    var r = pData; if (!r) return;
    $('pchips').innerHTML = r.kinds.filter(function(k){ return r.counts[k.key]; }).map(function(k){
      return '<a href="#" data-k="' + k.key + '" class="' + (!/^ok/.test(k.key) ? 'bad' : '') + (k.key == pKind ? ' on' : '') + '">' +
        esc(k.label) + ' (' + r.counts[k.key].toLocaleString() + ')</a>'; }).join('') ||
      '<span class="tm">No one is subscribed to this group in either place.</span>';
    Array.prototype.forEach.call(document.querySelectorAll('#pchips a'), function(a){
      a.onclick = function(e){ e.preventDefault(); pKind = a.getAttribute('data-k'); pRender(); }; });
    var k = r.kinds.filter(function(x){ return x.key == pKind; })[0];
    var bads = r.kinds.filter(function(x){ return !/^ok/.test(x.key) && r.counts[x.key]; }).length;
    $('phelp').innerHTML = k && k.help ? '<div class="khelp"' + (/^ok/.test(k.key) ? ' style="background:#f3f4f6;color:#374151"' : '') + '>' + esc(k.help) + '</div>' :
      (bads ? '' : '<div class="khelp" style="background:#f0fdf4;color:#166534">TouchPoint and PoweredByText agree on everyone in this group.</div>');
    var q = ($('pq').value || '').toLowerCase().replace(/[()\\- ]/g, '');
    var rows = r.rows.filter(function(x){
      if (x.kind != pKind) return false;
      if (!q) return true;
      var hay = (x.num + ' ' + x.pbtName + ' ' + x.people.map(function(p){ return p.name; }).join(' ')).toLowerCase().replace(/[()\\- ]/g, '');
      return hay.indexOf(q) >= 0;
    });
    var shown = rows.slice(0, LIMIT);
    $('ptab').innerHTML = !rows.length ? '<div class="empty" style="margin-top:30px">No one here.</div>' :
      '<table class="pt"><tr><th>Person in TouchPoint</th><th>Cell</th><th>TouchPoint</th><th>PoweredByText</th><th class="wide">Name in PoweredByText</th></tr>' +
      shown.map(function(x){
        var ppl = x.people.map(function(p){
          return '<a href="/Person2/' + p.pid + '" target="_blank">' + esc(p.name) + (p.age != null ? ' (' + p.age + ')' : '') + '</a>' +
            (!p.receive ? '<span class="pill p-optedout">texts off</span> ' : '');
        }).join('') || '<span class="tm">No one in TouchPoint has this number</span>';
        var tpCls = x.tp == 'sub' ? 'p-sub' : x.tp == 'opted' ? 'p-optedout' : 'p-none';
        var pbCls = !x.pbt ? 'p-none' : x.pbt.toLowerCase() == 'active' ? 'p-active' : 'p-optedout';
        return '<tr><td>' + ppl + (x.tp == '' && x.people.length ? ' <span class="tm">(on this number, not subscribed)</span>' : '') + '</td>' +
          '<td style="white-space:nowrap">' + fmtNum(x.num) + '</td>' +
          '<td><span class="pill ' + tpCls + '">' + TPL[x.tp] + '</span></td>' +
          '<td><span class="pill ' + pbCls + '">' + esc(x.pbt || 'Not listed') + '</span></td>' +
          '<td class="wide">' + esc(x.pbtName) + '</td></tr>';
      }).join('') + '</table>' +
      (rows.length > LIMIT ? '<div class="tm" style="margin-top:8px">Showing ' + LIMIT + ' of ' + rows.length.toLocaleString() + '. Search to narrow it.</div>' : '');
  }
  function showTab(t){
    Array.prototype.forEach.call(document.querySelectorAll('#tabs a'), function(a){ a.className = a.getAttribute('data-t') == t ? 'on' : ''; });
    $('pp').style.display = t == 'people' ? 'block' : 'none';
    $('wrap').style.display = t == 'people' ? 'none' : '';
    if (t == 'people') {
      if (!$('pgrp').options.length) {
        $('pgrp').innerHTML = groups.map(function(g){ return '<option value="' + g.id + '">' + esc(g.name) + '</option>'; }).join('');
        var pref = store('pgrp') || $('grp').value;
        if (pref && groups.some(function(g){ return String(g.id) == pref; })) $('pgrp').value = pref;
      }
      if (!pData || String(pData.gid) != $('pgrp').value) pLoad();
    }
  }
  Array.prototype.forEach.call(document.querySelectorAll('#tabs a'), function(a){
    a.onclick = function(e){ e.preventDefault(); if (groups.length) showTab(a.getAttribute('data-t')); }; });
  $('pgrp').onchange = function(){ store('pgrp', this.value); pKind = ''; pLoad(); };
  $('preload').onclick = pLoad;
  $('pq').oninput = pRender;

  // ------------------------------------------------ wiring
  Array.prototype.forEach.call(document.querySelectorAll('#seg a'), function(a){
    a.onclick = function(e){ e.preventDefault(); filt = a.getAttribute('data-f');
      Array.prototype.forEach.call(document.querySelectorAll('#seg a'), function(x){ x.className = x == a ? 'on' : ''; });
      renderList(); };
  });
  $('grp').onchange = function(){ store('grp', this.value); renderList(); };
  $('days').onchange = loadThreads;
  $('q').oninput = renderList;
  loadThreads();
  // Refresh pace: every 10 seconds while you are working in a conversation,
  // 15 while you are on the page, 60 after 5 minutes with no clicks or
  // typing. Coming back (a click after idling, or returning to the tab)
  // refreshes right away. Nothing refreshes while the tab is hidden.
  var lastAct = Date.now(), pollTimer = null;
  function pollDelay(){
    var idle = Date.now() - lastAct;
    if (idle > 5 * 60000) return 60000;
    return cur && idle < 2 * 60000 ? 10000 : 15000;
  }
  function poll(){
    clearTimeout(pollTimer);
    var app = document.getElementById('txapp');
    if (!document.hidden && !busy && $('pp').style.display != 'block' && !(app && app.className.indexOf('admin') >= 0)) {
      loadThreads();
      if (cur) openThread(cur.gid, cur.num, 'read');
    }
    pollTimer = setTimeout(poll, pollDelay());
  }
  ['click', 'keydown', 'touchstart'].forEach(function(ev){
    document.addEventListener(ev, function(){
      var wasIdle = Date.now() - lastAct > 5 * 60000;
      lastAct = Date.now();
      if (wasIdle) poll();
    }, true);
  });
  document.addEventListener('visibilitychange', function(){ if (!document.hidden) { lastAct = Date.now(); poll(); } });
  pollTimer = setTimeout(poll, pollDelay());
})();
</script>"""
        return (css + body.replace('__NEWHIDE__', '' if can_new else ' style="display:none"')
                + js.replace('__FIRST__', safe_json(first)).replace('__POLICY__', safe_json(send_policy())), True)


def admin_page():
    groups = pbt_groups()
    settings = load_settings()
    saved_groups = set(settings.get('groups') or [])
    configured = bool(model.Setting("PBTClientID", "")) and bool(model.Setting("PBTClientSecret", ""))
    installed = MB_START in mb_read()
    opts = ''.join(
        '<label class="g"><input type="checkbox" value="%d"%s> %s</label>' % (
            g['id'], '' if saved_groups and g['id'] not in saved_groups else ' checked',
            (g['name'] or '').replace('&', '&amp;').replace('<', '&lt;'))
        for g in groups)
    warn = '' if configured else \
        '<div class="warn">PBTClientID or PBTClientSecret is missing in Settings, so this cannot reach PoweredByText.</div>'
    # Settings an admin has to fill in before usage numbers mean anything.
    need = []
    if not settings.get('plan_credits') or not settings.get('plan_price'):
        need.append('your PBT plan (credits included and price)')
    if not settings.get('cycle_day'):
        need.append('the day your PBT billing period starts')
    if need:
        warn += ('<div class="warn" id="setupwarn"><b>Finish setting up.</b> Enter ' + ' and '.join(need) +
                 ' under Settings, so plan use, pace and dollars are right. '
                 '<button type="button" class="ghost" id="setupgo">Open Settings</button></div>')
    if settings.get('auto') and not installed:
        warn += ('<div class="warn">Morning updates are turned on, but the block in the MorningBatch script is '
                 'missing (someone may have edited it). Save the setting again to put it back.</div>')

    last = settings.get('last_auto') or {}
    if last:
        if last.get('stopped'):
            last_txt = 'Last morning run %s: %s' % (last.get('when', ''), last['stopped'])
        else:
            last_txt = 'Last morning run %s: %d changed, %d not confirmed, %d errors.' % (
                last.get('when', ''), last.get('changed', 0), last.get('unconfirmed', 0), last.get('errors', 0))
    else:
        last_txt = 'It has not run from the morning batch yet.'
    last_txt = last_txt.replace('&', '&amp;').replace('<', '&lt;')

    css = """
<style>
#pbt{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#222;max-width:1400px}
#pbt p.lede{color:#555;margin:0 0 14px}
#pbt .warn{background:#fdecea;border:1px solid #f5c2c0;padding:10px 14px;border-radius:6px;margin-bottom:12px}
#pbt .box{border:1px solid #ddd;border-radius:8px;padding:12px 14px;margin-bottom:14px;background:#fafafa}
#pbt .groups{display:flex;flex-wrap:wrap;gap:4px 18px;margin:8px 0}
#pbt .g{font-weight:normal;white-space:nowrap}
#pbt button{padding:7px 14px;border-radius:6px;border:1px solid #1e3a5f;background:#1e3a5f;color:#fff;cursor:pointer}
#pbt button.ghost{background:#fff;color:#1e3a5f}
#pbt button.go{background:#1d6b2c;border-color:#1d6b2c}
#pbt button:disabled{opacity:.5;cursor:default}
#pbt .cards{display:flex;flex-wrap:wrap;gap:10px;margin:12px 0}
#pbt .card{border:1px solid #ddd;border-radius:8px;padding:10px 14px;min-width:130px;cursor:pointer;background:#fff}
#pbt .card.on{outline:2px solid #1e3a5f}
#pbt .card b{display:block;font-size:22px}
#pbt .card span{font-size:12px;color:#666}
#pbt table{border-collapse:collapse;width:100%;font-size:13px}
#pbt th,#pbt td{border-bottom:1px solid #eee;padding:6px 8px;text-align:left;vertical-align:top}
#pbt th{background:#f3f3f3;position:sticky;top:0}
#pbt .tag{display:inline-block;padding:1px 7px;border-radius:10px;font-size:12px;font-weight:600}
#pbt .t-set{background:#e3f2e6;color:#1d6b2c}
#pbt .t-fix{background:#fff1d6;color:#8a5a00}
#pbt .t-same{background:#eee;color:#555}
#pbt .t-keep{background:#e8eefc;color:#2d4c9c}
#pbt .t-review{background:#fdecea;color:#a12a22}
#pbt .t-notfound{background:#f3f3f3;color:#888}
#pbt .r-changed{color:#1d6b2c;font-weight:600}
#pbt .r-unconfirmed{color:#8a5a00;font-weight:600}
#pbt .r-error,#pbt .r-stopped{color:#a12a22;font-weight:600}
#pbt .r-skipped{color:#666}
#pbt .who{color:#666;font-size:12px}
#pbt .kid{color:#a12a22}
#pbt #log,#pbt .status{font-size:12px;color:#666;margin-top:6px}
#pbt .bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:6px 0 10px}
#pbt input.filter{padding:6px 8px;border:1px solid #ccc;border-radius:6px;width:240px}
#pbt details.how{border:1px solid #ddd;border-radius:8px;padding:10px 14px;margin-bottom:14px;background:#fff}
#pbt details.how summary,#pbt details.box summary{font-weight:600;cursor:pointer}
#pbt details.how ol{margin:8px 0 4px;padding-left:22px}
#pbt details.how li{margin:5px 0}
#pbt details.how dl{display:grid;grid-template-columns:max-content 1fr;gap:4px 12px;margin:8px 0 0}
#pbt details.how dd{margin:0;color:#444}
#pbt details.how p,#pbt details.box p{margin:8px 0 0;color:#555}
#pbt .tabs{display:flex;gap:4px;margin-bottom:12px;border-bottom:1px solid #ddd}
#pbt .tabs a{padding:8px 14px;color:#1e3a5f;text-decoration:none;border:1px solid transparent;border-bottom:none;border-radius:6px 6px 0 0}
#pbt .tabs a.on{border-color:#ddd;background:#fff;font-weight:600;margin-bottom:-1px}
#pbt select,#pbt input[type=date]{padding:6px 8px;border:1px solid #ccc;border-radius:6px}
#pbt tr.send{cursor:pointer}
#pbt tr.send:hover td{background:#f6f9fc}
#pbt tr.detail td{background:#fbfbfb}
#pbt .late{color:#a12a22;font-weight:600}
#pbt .ontime{color:#1d6b2c}
#pbt .msg{max-width:420px;color:#444}
#pbt .num{text-align:right;font-variant-numeric:tabular-nums}
#pbt .note{font-size:12px;color:#666;margin:4px 0}
#pbt .k-delivered{color:#1d6b2c}
#pbt .k-failed{color:#a12a22;font-weight:600}
#pbt .k-other{color:#8a5a00}
#pbt table.mom th{font-size:12px;line-height:1.25;vertical-align:bottom}
#pbt table.mom th.num,#pbt table.mom td.num{white-space:nowrap;padding:5px 6px}
#pbt table.mom .g{position:sticky;left:0;background:#fff;z-index:1;min-width:150px}
#pbt table.mom thead .g{background:#f3f3f3;z-index:2}
#pbt table.mom .tag{font-size:10px;padding:0 5px;margin-top:2px}
#pbt table.mom td[data-dm]{cursor:pointer}
#pbt tr.dgrow{cursor:pointer}
#pbt tr.dgrow:hover td{background:#eef3fa}
#pbt tr.dgrow.on td{background:#e6edf6;font-weight:600}
#pbt table.mom td[data-dm]:hover{background:#eef3fa;text-decoration:underline}
#pbt table.mom td.picked{outline:2px solid #1e3a5f;outline-offset:-2px}
#pbt #mdrill .chip{display:inline-block;border:1px solid #d5dbe3;border-radius:4px;padding:0 6px;margin:2px 4px 2px 0;background:#f7f9fb;font-size:12px}
#pbt #mdrill .why{color:#555;font-size:12px;margin-top:3px}
#pbt[data-writes="off"] .writes-only,#pbt[data-writes="on"] .writes-off-only{display:none}
#pbt .t-multi{background:#fdecea;color:#9b2c1f}
</style>"""

    body = """
<div id="pbt" data-writes="__WRITES__">
  __WARN__
  <div class="tabs"><a href="#" data-tab="sync" class="on">Names</a><a href="#" data-tab="history">Change log</a><a href="#" data-tab="delivery">Delivery</a><a href="#" data-tab="usage">Usage</a><a href="#" data-tab="settings">Settings</a><a href="#" data-tab="check">Connection check</a></div>

  <div id="tab-sync">
  <p class="lede writes-only">Gives each texting number in PoweredByText the name of the person it belongs to in TouchPoint.
  Nothing changes until you pick rows and press a Change button, or turn on morning updates.</p>
  <p class="lede writes-off-only">Shows who each texting number in PoweredByText belongs to in TouchPoint, and which
  TouchPoint records need fixing. Read only.</p>
  <div class="box writes-off-only" style="border-left:4px solid #c98a00">
    <b>Changing names in PoweredByText is on hold.</b> PoweredByText does not yet let TouchPoint set a contact's name.
    Until it does, this tab is a report, and the data is still worth acting on:
    <ul style="margin:6px 0 0 18px">
      <li><b>On more than one record:</b> TouchPoint sends a group text once per person, so a parent's cell phone that is
        also on the children gets the same text two or three times, and each copy costs a credit. Take the number off
        the child records.</li>
      <li><b>Review:</b> no safe owner, usually a parent's phone that is only on a child's record.</li>
      <li><b>Not found:</b> texting numbers no one in TouchPoint has as a cell phone.</li>
    </ul>
  </div>
  <details class="how">
    <summary>Who gets the name when people share a number</summary>
    <p>Only people with this exact number in their Cell Phone field count. Deceased and archived records are skipped.
    The steps run in this order, and the first one that fits decides.</p>
    <ol>
      <li><b>Adults come before children.</b> When a parent's number was also typed on a child's record,
        the parent gets the name. A child counts as anyone listed as a child in the family or under 18.</li>
      <li><b>One adult:</b> that adult gets the name.</li>
      <li><b>Two or more adults in the same family:</b> the head of household. If no one is marked head,
        the oldest primary adult.</li>
      <li><b>Adults in different families:</b> no guess. The number is listed under <b>Review</b>.</li>
      <li><b>Only children:</b> if exactly one is 13 or older, that teen gets the name.
        Otherwise the number is listed under <b>Review</b>, since it is almost always a parent's phone
        that is missing from the parent's record.</li>
    </ol>
    <p>The name used is the nickname (Goes By) when there is one, then the last name.</p>
    <p class="writes-only"><b>What happens to the name already in PoweredByText</b></p>
    <dl class="writes-only">
      <dt><span class="tag t-set">Will set</span></dt><dd>There is no name, so the chosen person's name is added.</dd>
      <dt><span class="tag t-fix">Will fix</span></dt><dd>The name belongs to someone else on this number, usually the child, and is replaced.</dd>
      <dt><span class="tag t-same">Already right</span></dt><dd>Nothing changes.</dd>
      <dt><span class="tag t-keep">Keep</span></dt><dd>The name matches no one on this number, so someone typed it in PoweredByText. It is never overwritten.</dd>
      <dt><span class="tag t-review">Review</span></dt><dd>Nothing changes. Fix the records in TouchPoint (put the number on the right person, or take it off the wrong one) and the next run picks it up.</dd>
      <dt><span class="tag t-notfound">Not found</span></dt><dd>No one in TouchPoint has this cell phone. Nothing changes.</dd>
    </dl>
    <p>To change who gets a number, change the records in TouchPoint, not this page.
    Each change also stores the person's TouchPoint ID on the PoweredByText contact.</p>
  </details>

  <div class="box">
    <div><b>Groups</b>
      <a href="#" id="all">all</a> &middot; <a href="#" id="none">none</a></div>
    <div class="groups" id="groups">__OPTS__</div>
    <button id="run">Run preview</button>
    <button id="csv" class="ghost" disabled>Download CSV</button>
    <div id="nprog" style="display:none;margin:6px 0 10px">
      <div style="height:8px;background:#e6e9ee;border-radius:4px;overflow:hidden"><div id="nbar" style="height:100%;width:0;background:#1e3a5f;transition:width .3s"></div></div>
      <div class="note" id="nprogtxt"></div>
    </div>
    <div id="log"></div>
  </div>

  <details class="box writes-only">
    <summary>Morning updates</summary>
    <p><label><input type="checkbox" id="auto"__AUTO__> Update names every morning for the groups checked above</label></p>
    <p>Each morning it makes the <b>Will set</b> and <b>Will fix</b> changes, the same as pressing Change here.
    If more than __AUTOMAX__ names would change in one morning, it changes nothing and says so in the change log,
    so a mistake in TouchPoint can't rename everyone. Do the first big catch-up from this page.</p>
    <p>Saving adds a small block to the MorningBatch script, or removes it when turned off. Nothing else in that script changes.</p>
    <p><button id="saveauto">Save</button> <span class="status" id="autostatus">__INSTALLED__ __LAST__</span></p>
  </details>

  <div class="cards" id="cards"></div>
  <div class="bar">
    <input class="filter" id="nq" placeholder="Filter by name or number">
    <button id="applysel" class="go writes-only" disabled>Change selected</button>
    <button id="applyall" class="ghost writes-only" disabled>Change all Will set and Will fix</button>
    <span class="status" id="applystatus"></span>
  </div>
  <table><thead><tr>
    <th><input type="checkbox" id="selall" class="writes-only" title="Select every changeable row shown"></th>
    <th>Number</th><th>Groups</th><th>Name in PBT now</th><th id="hstatus">Sync would</th><th id="hname">New name</th>
    <th>Why</th><th>People in TouchPoint with this number</th>
  </tr></thead><tbody id="rows"><tr><td colspan="8" class="who" style="padding:14px 8px">Pick the groups above and press <b>Run preview</b> to see who each texting number belongs to in TouchPoint and which records need fixing. It only reads.</td></tr></tbody></table>
  </div>

  <div id="tab-delivery" style="display:none">
    <p class="note">When each group text actually went out, how long it took to reach everyone, and how it went.
    Texts sent from TouchPoint show the time TouchPoint was asked to send them. PBT does not report the
    scheduled time of a text scheduled inside PBT, only when it went out. Read only.</p>
    <div class="bar">
      <select id="dgroup"><option value="0">All groups</option>__GROUPOPTS__</select>
      <input type="date" id="dfrom"> to <input type="date" id="dto">
      <button id="dload">Load</button>
      <label><input type="checkbox" id="dsingle" checked> Hide texts to one person</label>
      <input class="filter" id="dnum" placeholder="Did this number get it? (10 digits)">
      <span class="status" id="dstatus"></span>
    </div>
    <div id="dprog" style="display:none;margin:6px 0 10px">
      <div style="height:8px;background:#e6e9ee;border-radius:4px;overflow:hidden"><div id="dbar" style="height:100%;width:0;background:#1e3a5f;transition:width .3s"></div></div>
      <div class="note" id="dprogtxt"></div>
    </div>
    <div id="dnotes" class="note"></div>
    <div id="dout"><p class="note">Pick a group (or all of them) and the dates, then press <b>Load</b> to see when each group text went out and how it went.</p></div>
  </div>

  <div id="tab-usage" style="display:none">
    <div class="cards" id="ukpi"></div>
    <div class="note" id="ukpinote"></div>
    <details class="box" id="mom" open>
      <summary>Billing period over period, every group</summary>
      <div class="bar">
        <select id="mmonths"><option value="6" selected>Last 6 periods</option><option value="12">Last 12 periods</option><option value="24">Last 24 periods</option></select>
        <button id="mbuild">Show</button>
        <button id="mrefresh" class="ghost" title="Read what is new in PoweredByText since the last refresh">Refresh</button>
        <select id="mmetric"><option value="credits">Credits</option><option value="waste">Waste</option>
          <option value="failedCredits">Not delivered</option><option value="extra">Emoji extra</option>
          <option value="dupCredits">Same person twice</option><option value="fallbackCredits">Fallback texts</option>
          <option value="sent">Texts sent</option>
          <option value="unknown">No final status (texts)</option>
          <option value="notTp">Not in TouchPoint (texts)</option></select>
        <button id="mcsv" class="ghost">Download CSV</button>
        <label title="Normally a finished month is counted once and saved. Tick this to count every month again."><input type="checkbox" id="mrecount"> Recount saved months</label>
        <span class="status" id="mstatus"></span>
      </div>
      <div id="mprog" style="display:none;margin:6px 0 10px">
        <div style="height:8px;background:#e6e9ee;border-radius:4px;overflow:hidden"><div id="mbar" style="height:100%;width:0;background:#1e3a5f;transition:width .3s"></div></div>
        <div class="note" id="mprogtxt"></div>
      </div>
      <div id="mout"></div>
      <div id="mdrill"></div>
    </details>
    <details class="box" id="urange">
      <summary>Look closer at a date range</summary>
    <p class="note">Estimated credits and where they go, counted from PBT's message report. Read only.</p>
    <div class="bar">
      <select id="ugroup"><option value="0">All groups</option>__GROUPOPTS__</select>
      <input type="date" id="ufrom"> to <input type="date" id="uto">
      <button id="uload">Count</button>
      <span class="status" id="ucstatus">Plan, cost and billing cycle are under Settings.</span>
    </div>
    <div id="ucprog" style="display:none;margin:6px 0 10px">
      <div style="height:8px;background:#e6e9ee;border-radius:4px;overflow:hidden"><div id="ucbar" style="height:100%;width:0;background:#1e3a5f;transition:width .3s"></div></div>
      <div class="note" id="ucprogtxt"></div>
    </div>
    <div id="unotes" class="note"></div>
    <div id="uout"><p class="note">Pick a group and dates, then press <b>Count</b> for the credits, waste, costliest texts and repeats in that range.</p></div>
    </details>
  </div>

  <div id="tab-settings" style="display:none">
    <div class="box">
      <p><b>Plan and billing</b> <span class="who">From your PoweredByText invoice.</span></p>
      <p><label>Your plan includes <input type="text" id="uplan" size="7" placeholder="30000" value="__PLANCREDITS__"> credits
        each billing period, for $<input type="text" id="uprice" size="6" placeholder="950" value="__PLANPRICE__"></label></p>
      <p><label>Each credit over the plan costs $<input type="text" id="uover" size="6" placeholder="0.04" value="__OVERAGE__"></label></p>
      <p><label>The billing period starts on day <input type="text" id="ucycle" size="2" value="__CYCLEDAY__"> of the month</label>
        <span class="who">(18 means the 18th through the 17th)</span></p>
      <p class="who" id="ucosthint"></p>
      <input type="hidden" id="ucost" value="">
    </div>
    <div class="box">
      <p><b>Who sees what</b></p>
      <p><label><input type="checkbox" id="setadminall"__ADMINALL__> Admins see every texting group&rsquo;s conversations</label>
        <span class="who">Off: everyone, admins included, sees only the groups they are listed in on SMS Management.
        On: admins can read every group, but can send only from groups they are members of.</span></p>
    </div>
    <div class="box">
      <p><b>What staff can send</b></p>
      <p><label>Emoji <select id="setemoji" data-v="__EMOJIPOL__">
        <option value="allow">Allow</option>
        <option value="replies">Allow in replies, warn in group texts</option>
        <option value="none">Do not allow</option></select></label>
        <span class="who">One emoji in a plain text can add a credit for every person it goes to. In a picture message it is free.</span></p>
      <p><label>Attachments <select id="setattach" data-v="__ATTACHPOL__">
        <option value="off">Off</option>
        <option value="pictures">Pictures only</option>
        <option value="pictures_pdf">Pictures and PDFs</option>
        <option value="all">Pictures, PDFs, small video and audio</option></select></label>
        <span class="who">Always 3 credits. Photos are shrunk to fit the 950 KB limit; anything bigger is refused with a tip to send a link.</span></p>
    </div>
    <p><button id="ucostsave">Save settings</button> <span class="status" id="ustatus"></span></p>
  </div>

  <div id="tab-check" style="display:none">
    <p class="note">Checks each step between TouchPoint and PoweredByText, so you can tell where a problem is. It only reads.</p>
    <div class="bar"><button id="ckrun">Run the check</button>
      <button id="cklinks" class="ghost" title="Adds one short link to the first group's short link list in PoweredByText">Test short links</button>
      <span class="status" id="ckstatus"></span></div>
    <div id="ckprog" style="display:none;margin:6px 0 10px">
      <div style="height:8px;background:#e6e9ee;border-radius:4px;overflow:hidden"><div id="ckbar" style="height:100%;width:0;background:#1e3a5f;transition:width .3s"></div></div>
      <div class="note" id="ckprogtxt"></div>
    </div>
    <div id="ckout"><p class="note">Press <b>Run the check</b>. It takes a few seconds and changes nothing.</p></div>
    <details class="box" style="margin-top:14px"><summary>Advanced</summary>
      <p class="note">Compare PoweredByText's raw message report with what TouchPoint stored: every field, incoming texts matched,
      and how far apart the two clocks are. For tracking down a problem with PoweredByText support.
      <a href="?view=pbttest" target="_blank">Open the report comparison</a></p>
    </details>
  </div>

  <div id="tab-history" style="display:none">
    <div class="bar">
      <input class="filter" id="hq" placeholder="Filter the log">
      <button id="hcsv" class="ghost">Download CSV</button>
      <span class="status" id="hstatus"></span>
    </div>
    <table><thead><tr>
      <th>When</th><th>By</th><th>Number</th><th>Group</th><th>Before</th><th>After</th><th>Result</th><th>Note</th>
    </tr></thead><tbody id="hrows"></tbody></table>
  </div>
</div>"""
    gopts = ''.join('<option value="%d">%s</option>' % (g['id'], (g['name'] or '').replace('&', '&amp;').replace('<', '&lt;'))
                    for g in groups)
    cc = settings.get('credit_cost')
    body = (body.replace('__WARN__', warn)
                .replace('__CREDITCOST__', ('%g' % cc) if cc else '')
                .replace('__PLANCREDITS__', str(settings.get('plan_credits') or ''))
                .replace('__PLANPRICE__', ('%g' % settings['plan_price']) if settings.get('plan_price') else '')
                .replace('__OVERAGE__', ('%g' % settings['overage_cost']) if settings.get('overage_cost') else '')
                .replace('__CYCLEDAY__', str(settings.get('cycle_day') or 1))
                .replace('__EMOJIPOL__', send_policy()['emoji'])
                .replace('__ADMINALL__', ' checked' if settings.get('admin_all_groups') else '')
                .replace('__ATTACHPOL__', send_policy()['attach'])
                .replace('__GROUPOPTS__', gopts)
                .replace('__OPTS__', opts or '<i>No PoweredByText groups found.</i>')
                .replace('__AUTO__', ' checked' if settings.get('auto') else '')
                .replace('__WRITES__', 'on' if NAME_WRITES else 'off')
                .replace('__AUTOMAX__', str(AUTO_MAX))
                .replace('__INSTALLED__', 'In the morning batch.' if installed else 'Not in the morning batch.')
                .replace('__LAST__', last_txt))

    js = """
<script>
(function(){
  var LABEL = {set:'Will set', fix:'Will fix', same:'Already right', keep:'Keep', review:'Review', notfound:'Not found'};
  var HELP = {set:'PBT has no name', fix:'PBT has another family member\\'s name', same:'Nothing to change',
              keep:'Typed in PBT; never overwritten', review:'Clean up in TouchPoint first', notfound:'Not in TouchPoint'};
  var RESULT = {changed:'Changed', unconfirmed:'Not confirmed', error:'Error', skipped:'Skipped', stopped:'Stopped'};
  var ORDER = ['set','fix','review','keep','same','notfound'];
  var WRITES = document.getElementById('pbt').getAttribute('data-writes') == 'on';
  if (!WRITES) {
    LABEL.set = 'No name in PBT'; LABEL.fix = 'Wrong name in PBT';
    HELP.set = 'TouchPoint knows who it is'; HELP.fix = 'Shows another family member';
    LABEL.multi = 'On more than one record'; HELP.multi = 'Gets each group text once per person';
    document.getElementById('hstatus').textContent = 'Status';
    document.getElementById('hname').textContent = 'Belongs to';
  }
  function multi(r){ return r.cands && r.cands.length > 1; }
  var merged = {}, filter = null, sel = {}, busy = false, hist = [];
  var base = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
  function $(id){ return document.getElementById(id); }
  function esc(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
  function fmt(n){ return n && n.length == 10 ? '(' + n.substr(0,3) + ') ' + n.substr(3,3) + '-' + n.substr(6) : n; }
  function post(data, cb){
    var x = new XMLHttpRequest();
    x.open('POST', base, true);
    x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    x.onload = function(){ var r; try { r = JSON.parse(x.responseText); } catch(e) {
      r = {success:false, message:'Bad response: ' + x.responseText.substr(0,120)}; } cb(r); };
    x.onerror = function(){ cb({success:false, message:'Network error'}); };
    x.send(Object.keys(data).map(function(k){ return k + '=' + encodeURIComponent(data[k]); }).join('&'));
  }
  function csv(name, head, rows){
    var lines = [head].concat(rows).map(function(cols){ return cols.map(function(v){
      return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"'; }).join(','); });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([lines.join('\\r\\n')], {type:'text/csv'}));
    a.download = name; document.body.appendChild(a); a.click(); a.remove();
  }
  function boxes(){ return Array.prototype.slice.call(document.querySelectorAll('#groups input')); }
  function changeable(r){ return WRITES && (r.action == 'set' || r.action == 'fix') && r.result != 'changed'; }
  $('all').onclick = function(e){ e.preventDefault(); boxes().forEach(function(b){ b.checked = true; }); };
  $('none').onclick = function(e){ e.preventDefault(); boxes().forEach(function(b){ b.checked = false; }); };

  Array.prototype.forEach.call(document.querySelectorAll('#pbt .tabs a'), function(a){
    a.onclick = function(e){
      e.preventDefault();
      var t = this.getAttribute('data-tab');
      Array.prototype.forEach.call(document.querySelectorAll('#pbt .tabs a'), function(x){ x.className = x == a ? 'on' : ''; });
      $('tab-sync').style.display = t == 'sync' ? '' : 'none';
      $('tab-history').style.display = t == 'history' ? '' : 'none';
      $('tab-delivery').style.display = t == 'delivery' ? '' : 'none';
      $('tab-usage').style.display = t == 'usage' ? '' : 'none';
      $('tab-settings').style.display = t == 'settings' ? '' : 'none';
      $('tab-check').style.display = t == 'check' ? '' : 'none';
      if (t == 'usage' && !momData && !$('mbuild').disabled) build('saved');
      if (t == 'history') loadHistory();
    };
  });

  // ---- Delivery and Usage (read only)
  function day(d){ return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  (function(){
    var t = new Date(), w = new Date(t.getTime() - 6 * 864e5), m = new Date(t.getFullYear(), t.getMonth(), 1);
    $('dfrom').value = day(w); $('dto').value = day(t);
    var cd = parseInt($('ucycle').value, 10) || 1;
    var ps = new Date(t.getFullYear(), t.getMonth(), cd); if (t.getDate() < cd) ps.setMonth(ps.getMonth() - 1);
    $('ufrom').value = day(ps); $('uto').value = day(t);
  })();
  function mins(m){ if (m < 60) return m + 'm'; return Math.floor(m / 60) + 'h ' + (m % 60) + 'm'; }
  function n(x){ return (x || 0).toLocaleString(); }
  var sends = [], allSends = [];
  // Which group's sends are listed under the summary. null = none yet
  // (click a group), '*' = all of them.
  var dPick = null;
  function pickSends(){
    var hide = $('dsingle').checked;
    sends = allSends.filter(function(s){ return !hide || s.n > 1; });
  }
  $('dsingle').onchange = function(){ pickSends(); drawSends(); };
  // The same bar as Refresh, for a page that waits on one request. The server
  // cannot report steps, so the label says what is being read and for how
  // long; the bar eases toward the end by the size of the range and fills
  // when the answer arrives.
  function waitBar(k, what, days, btn){
    var t0 = Date.now(), expect = Math.max(4, days * 1.2);
    $(k + 'prog').style.display = ''; $(k + 'bar').style.width = '0';
    if (btn) $(btn).disabled = true;
    function tick(){
      var el = (Date.now() - t0) / 1000;
      $(k + 'bar').style.width = Math.round(92 * (1 - Math.exp(-el / expect))) + '%';
      $(k + 'progtxt').textContent = what + ', ' + Math.round(el) + 's so far. Longer ranges take longer.';
    }
    tick(); var timer = setInterval(tick, 1000);
    return function(){ clearInterval(timer); $(k + 'bar').style.width = '100%'; if (btn) $(btn).disabled = false;
      setTimeout(function(){ $(k + 'prog').style.display = 'none'; }, 300); return Math.round((Date.now() - t0) / 1000); };
  }
  function rangeText(a, b){ return a == b ? md(a) : md(a) + ' to ' + md(b); }
  function rangeDays(a, b){ return Math.max(1, Math.round((new Date(b) - new Date(a)) / 864e5) + 1); }
  $('dload').onclick = function(){
    $('dstatus').textContent = ''; $('dout').innerHTML = ''; $('dnotes').textContent = '';
    var a = $('dfrom').value, b = $('dto').value;
    var done = waitBar('d', 'Reading ' + rangeText(a, b) + ' from PoweredByText', rangeDays(a, b), 'dload');
    post({action:'trace', gid:$('dgroup').value, from:a, to:b}, function(r){
      var took = done();
      if (!r.success) { $('dstatus').textContent = r.message; return; }
      allSends = r.sends || [];
      dPick = $('dgroup').value != '0' ? '*' : null;
      pickSends();
      $('dstatus').textContent = allSends.length + ' sends from ' + n(r.rowCount) + ' report rows, in ' + took + 's.';
      $('dnotes').textContent = (r.notes || []).join(' ');
      drawSends();
    });
  };
  function secs(x){ if (x < 60) return x + 's'; return mins(Math.round(x / 60)); }
  function pct(a, b){ return b ? Math.round(1000 * a / b) / 10 + '%' : '-'; }
  function hm(ts){ return ts ? ts.substr(5, 11) : ''; }
  function summary(){
    var tot = {sends: sends.length, people: 0, del: 0, fail: 0, other: 0}, tp = [], spreads = [], groups = {};
    sends.forEach(function(s){
      tot.people += s.n; tot.del += s.counts.delivered; tot.fail += s.counts.failed; tot.other += s.counts.other;
      spreads.push(s.spreadMin);
      if (s.tp) tp.push(s.tp.lateMin);
      var g = groups[s.group] = groups[s.group] || {sends: 0, people: 0, del: 0, fail: 0, waits: [], spreads: []};
      g.sends++; g.people += s.n; g.del += s.counts.delivered; g.fail += s.counts.failed; g.spreads.push(s.spreadMin);
      if (s.tp) g.waits.push(s.tp.lateMin);
    });
    function avg(a){ return a.length ? Math.round(a.reduce(function(x, y){ return x + y; }, 0) / a.length) : null; }
    function mx(a){ return a.length ? Math.max.apply(null, a) : null; }
    var lateN = tp.filter(function(m){ return m > 10; }).length;
    var h = '<div class="cards">'
      + '<div class="card"><b>' + n(tot.sends) + '</b><span>group texts</span></div>'
      + '<div class="card"><b>' + n(tot.people) + '</b><span>people texted</span></div>'
      + '<div class="card"><b>' + pct(tot.del, tot.people) + '</b><span>delivered</span></div>'
      + '<div class="card"><b>' + n(tot.fail) + '</b><span>failed (' + pct(tot.fail, tot.people) + ')</span></div>'
      + '<div class="card"><b>' + (tp.length ? mins(Math.max(0, avg(tp))) : '-') + '</b><span>avg wait, scheduled to out (' + tp.length + ' from TouchPoint)</span></div>'
      + '<div class="card"><b>' + (tp.length ? mins(Math.max(0, mx(tp))) : '-') + '</b><span>longest wait' + (lateN ? ', ' + lateN + ' over 10m' : '') + '</span></div>'
      + '<div class="card"><b>' + (spreads.length ? mins(avg(spreads)) : '-') + '</b><span>avg time to reach everyone</span></div>'
      + '</div>';
    var names = Object.keys(groups).sort();
    if (names.length > 1) {
      h += '<table style="margin-bottom:14px"><thead><tr><th>Group</th><th class="num">Texts</th><th class="num">People</th><th class="num">Delivered</th>'
        + '<th class="num">Failed</th><th class="num">Avg wait</th><th class="num">Longest wait</th><th class="num">Avg to reach all</th></tr></thead><tbody>'
        + names.map(function(k){ var g = groups[k];
            return '<tr class="dgrow' + (dPick == k ? ' on' : '') + '" data-g="' + esc(k) + '" title="Show the texts from this group"><td>' + esc(k) + '</td><td class="num">' + n(g.sends) + '</td><td class="num">' + n(g.people) + '</td><td class="num">'
              + pct(g.del, g.people) + '</td><td class="num' + (g.fail ? ' k-failed' : '') + '">' + n(g.fail) + '</td><td class="num">'
              + (g.waits.length ? mins(Math.max(0, avg(g.waits))) : '-') + '</td><td class="num">' + (g.waits.length ? mins(Math.max(0, mx(g.waits))) : '-')
              + '</td><td class="num">' + mins(avg(g.spreads)) + '</td></tr>'; }).join('') + '</tbody></table>';
    }
    var lateRes = 0; sends.forEach(function(s){ if (s.late) lateRes += s.late.n; });
    if (lateRes) h += '<p class="note">' + n(lateRes) + ' results came back hours or days after their text went out (a phone that was off; '
      + 'the carrier keeps trying for about two days). They are shown under their send, not as a slow send.</p>';
    h += '<p class="note">Wait is from the time TouchPoint was asked to send a text to the first recipient time PBT recorded; '
      + 'texts scheduled inside PBT have no scheduled time to measure from. PBT records one time per recipient, '
      + 'with the delivery status for that text beside it.</p>';
    return h;
  }
  function drawSends(){
    if (!sends.length) { $('dout').innerHTML = '<p class="note">No group texts in this range.</p>'; return; }
    var gnames = {}; sends.forEach(function(s){ gnames[s.group] = 1; });
    var one = Object.keys(gnames).length < 2, list = one || dPick == '*' ? sends : sends.filter(function(s){ return s.group == dPick; });
    var h = summary();
    if (!one && !dPick) {
      $('dout').innerHTML = h + '<p class="note"><b>Click a group above</b> to list each text it sent, or <a href="#" id="dall">show all '
        + n(sends.length) + ' texts</a>.</p>';
      wireGroups(); return;
    }
    if (!one) h += '<p class="note">' + (dPick == '*' ? 'All groups' : '<b>' + esc(dPick) + '</b>') + ': ' + n(list.length) + ' text' + (list.length == 1 ? '' : 's')
      + '. ' + (dPick == '*' ? '' : '<a href="#" id="dall">Show all groups</a>. ') + 'Click a text to see each person.</p>';
    h += '<table><thead><tr><th>Created</th><th>Scheduled for</th><th>First out</th><th>Reached all by</th>'
      + '<th>Group</th><th>How</th><th class="num">To</th><th class="num">Delivered</th><th class="num">Failed</th><th class="num">Other</th><th>Message</th></tr></thead><tbody>';
    list.forEach(function(s, i){
      var created = '<span class="who">in PBT</span>', due = '<span class="who">not in TouchPoint</span>', wait = '';
      if (s.tp) {
        created = esc(hm(s.tp.created)) + '<div class="who">' + esc(s.tp.sender || '') + (s.tp.match == 'text' ? ', matched by text' : '') + '</div>';
        due = s.tp.scheduled ? esc(hm(s.tp.due)) : '<span class="who">send now</span>';
        var late = s.tp.lateMin;
        wait = '<div class="' + (late > 10 ? 'late' : 'ontime') + '">' + (late > 10 ? mins(late) + ' late' : 'on time') + '</div>';
      }
      h += '<tr class="send" data-i="' + i + '"><td>' + created + '</td><td>' + due + '</td><td>' + esc(hm(s.first)) + wait
        + '</td><td>' + esc(hm(s.last)) + '<div class="who">' + (s.spreadMin ? mins(s.spreadMin) + ' after first' : 'all within 1m')
        + ', avg ' + secs(s.avgSec) + '</div>' + lateNote(s) + '</td><td>' + esc(s.group) + '</td><td>' + esc(s.typeLabel) + (s.mms ? ' <span class="tag t-same">picture</span>' : '')
        + '</td><td class="num">' + n(s.n) + '</td><td class="num k-delivered">' + n(s.counts.delivered)
        + '</td><td class="num' + (s.counts.failed ? ' k-failed' : '') + '">' + n(s.counts.failed)
        + '</td><td class="num">' + n(s.counts.other) + '</td><td class="msg">' + esc((s.message || '').substr(0, 140)) + '</td></tr>';
    });
    $('dout').innerHTML = h + '</tbody></table>';
    wireGroups();
    Array.prototype.forEach.call(document.querySelectorAll('#dout tr.send'), function(tr){
      tr.onclick = function(){
        var nx = tr.nextSibling;
        if (nx && nx.className == 'detail') { nx.remove(); return; }
        var s = list[+tr.getAttribute('data-i')];
        var rows = s.people.slice().sort(function(a, b){ return a[4] == b[4] ? (a[2] < b[2] ? -1 : 1) : (a[4] == 'failed' ? -1 : b[4] == 'failed' ? 1 : 0); });
        var d = '<table><thead><tr><th>Number</th><th>Whose number</th><th>Time</th><th>Status</th></tr></thead><tbody>'
          + rows.map(function(p){ return '<tr><td>' + fmt(p[0]) + '</td><td>' + who(p) + '</td><td>' + esc(p[2])
            + '</td><td class="k-' + p[4] + '">' + esc(p[3]) + '</td></tr>'; }).join('') + '</tbody></table>';
        if (s.latePeople && s.latePeople.length) {
          d += '<p class="note" style="margin-top:8px"><b>Late results.</b> The carrier answered for these numbers well after the send, '
            + 'usually because the phone was off or out of service. ' + (s.late.again ? s.late.again + ' of them were already in the send above, reported again. ' : '')
            + 'They are not counted in how fast it went out.</p>'
            + '<table><thead><tr><th>Number</th><th>Whose number</th><th>Answered</th><th>Status</th></tr></thead><tbody>'
            + s.latePeople.map(function(p){ return '<tr><td>' + fmt(p[0]) + '</td><td>' + who(p) + '</td><td>' + esc(p[2])
              + '</td><td class="k-' + p[4] + '">' + esc(p[3]) + '</td></tr>'; }).join('') + '</tbody></table>';
        }
        var row = document.createElement('tr'); row.className = 'detail';
        row.innerHTML = '<td colspan="11">' + d + '</td>';
        tr.parentNode.insertBefore(row, tr.nextSibling);
      };
    });
  }
  function wireGroups(){
    Array.prototype.forEach.call(document.querySelectorAll('#dout tr.dgrow'), function(tr){
      tr.onclick = function(){ dPick = tr.getAttribute('data-g'); drawSends(); };
    });
    if ($('dall')) $('dall').onclick = function(e){ e.preventDefault(); dPick = '*'; drawSends(); };
  }
  function lateNote(s){
    if (!s.late) return '';
    var L = s.late;
    return '<div class="who" title="PBT stamps each person with the time the carrier answered. For a phone that is off, '
      + 'the carrier keeps trying for about two days, then reports delivered or failed. These are not a second send '
      + 'and are left out of the times above.">+ ' + L.n + ' late result' + (L.n == 1 ? '' : 's') + ' up to '
      + (L.hours >= 24 ? Math.round(L.hours / 24 * 10) / 10 + ' days' : L.hours + 'h') + ' later (' + L.delivered + ' delivered, '
      + L.failed + ' failed): phone off, carrier kept trying</div>';
  }
  function who(p){
    var nm = p[5] ? '<a href="/Person2/' + p[5] + '" target="_blank">' + esc(p[1]) + '</a>' : esc(p[1]);
    return nm + (p[6] ? '<div class="who">' + esc(p[6]) + '</div>' : '');
  }
  $('dnum').oninput = function(){
    var q = this.value.replace(/[^0-9]/g, '');
    if (q.length < 10) { if (!q) drawSends(); return; }
    q = q.slice(-10);
    var hits = [];
    sends.forEach(function(s){ s.people.concat(s.latePeople || []).forEach(function(p){ if (p[0] == q) hits.push([s, p]); }); });
    $('dout').innerHTML = hits.length
      ? '<p class="note">' + fmt(q) + ': ' + who(hits[0][1]) + '</p>'
        + '<table><thead><tr><th>Went out</th><th>Group</th><th>Status</th><th>Message</th></tr></thead><tbody>'
        + hits.map(function(h){ return '<tr><td>' + esc(h[0].first.substr(5, 6) + h[1][2]) + '</td><td>' + esc(h[0].group)
          + '</td><td class="k-' + h[1][4] + '">' + esc(h[1][3]) + '</td><td class="msg">' + esc(h[0].message.substr(0, 140)) + '</td></tr>'; }).join('')
        + '</tbody></table>'
      : '<p class="note">' + fmt(q) + ' got none of the group texts loaded above.</p>';
  };
  var lastUsage = null;
  function numv(id){ var v = parseFloat(($(id).value || '').replace(/[^0-9.]/g, '')); return isNaN(v) || v <= 0 ? 0 : v; }
  function planRate(){ var c = numv('uplan'), p = numv('uprice'); return c && p ? p / c : 0; }
  function rate(){ return numv('ucost') || planRate(); }
  function syncCostHint(){
    var pr = planRate();
    $('ucosthint').textContent = pr ? 'Each credit in your plan works out to $' + pr.toFixed(4) + ' ($' + numv('uprice')
      + ' divided by ' + numv('uplan').toLocaleString() + ' credits). Waste is shown in dollars at that price.' : 'Fill in your plan to see costs and waste in dollars.';
  }
  syncCostHint();
  function money(cr){ var r = rate(); if (!r || cr == null) return ''; var v = cr * r;
    return '$' + v.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}); }
  function crm(cr){ var m = money(cr); return n(cr) + (m ? '<div class="who">' + m + '</div>' : ''); }
  $('ucostsave').onclick = function(){
    var before = momData ? cycleDay() : null;
    post({action:'save_cost', cost:$('ucost').value.replace(/[^0-9.]/g, ''), cycle_day:$('ucycle').value.replace(/[^0-9]/g, ''),
          plan_credits:$('uplan').value.replace(/[^0-9]/g, ''), plan_price:$('uprice').value.replace(/[^0-9.]/g, ''),
          overage_cost:$('uover').value.replace(/[^0-9.]/g, ''),
          emoji_policy:$('setemoji').value, attach_policy:$('setattach').value, admin_all:$('setadminall').checked ? '1' : '0'}, function(r){
      $('ustatus').textContent = r.success ? 'Saved.' : r.message;
      if (r.success && $('setupwarn') && r.planCredits && r.planPrice && $('ucycle').value) $('setupwarn').style.display = 'none';
      if (r.success && momData) $('mstatus').textContent = 'Press Show to count by the saved billing cycle.';
      if (r.success && lastUsage) drawUsage(lastUsage);
      if (r.success) drawMom();
    });
  };
  if ($('setupgo')) $('setupgo').onclick = function(){
    var a = document.querySelector('#txnav a[data-p="settings"]');
    if (a) a.click(); else document.querySelector('#pbt .tabs a[data-tab="settings"]').click();
  };
  // ---- Connection check
  var CK = {ok: ['&#10003;', 'k-delivered'], warn: ['!', 'late'], bad: ['&#10007;', 'k-failed'], info: ['i', 'who']};
  function ckRow(c){
    var m = CK[c.state] || CK.info;
    return '<tr><td class="' + m[1] + '" style="font-weight:700;text-align:center;width:28px">' + m[0] + '</td><td style="white-space:nowrap"><b>'
      + esc(c.name) + '</b></td><td>' + esc(c.detail) + '</td></tr>';
  }
  $('ckrun').onclick = function(){
    $('ckout').innerHTML = ''; $('ckstatus').textContent = '';
    var done = waitBar('ck', 'Checking TouchPoint and PoweredByText', 6, 'ckrun');
    post({action:'pbt_check'}, function(r){
      var took = done();
      if (!r.success) { $('ckstatus').textContent = r.message; return; }
      var bad = r.checks.filter(function(c){ return c.state == 'bad'; }).length, warn = r.checks.filter(function(c){ return c.state == 'warn'; }).length;
      $('ckstatus').textContent = (bad ? bad + ' problem' + (bad == 1 ? '' : 's') : warn ? warn + ' to look at' : 'Everything checks out') + ', in ' + took + 's.';
      $('ckout').innerHTML = '<table>' + r.checks.map(ckRow).join('') + '</table>';
    });
  };
  $('cklinks').onclick = function(){
    var done = waitBar('ck', 'Asking PoweredByText to shorten a test link', 2, 'cklinks');
    post({action:'pbt_check_links'}, function(r){
      done();
      var row = ckRow({name: 'Short links', state: r.success ? r.state : 'bad', detail: r.success ? r.detail : r.message});
      var t = $('ckout').querySelector('table');
      if (t) t.insertAdjacentHTML('beforeend', row); else $('ckout').innerHTML = '<table>' + row + '</table>';
    });
  };
  $('setemoji').value = $('setemoji').getAttribute('data-v') || 'replies';
  $('setattach').value = $('setattach').getAttribute('data-v') || 'pictures_pdf';
  $('ucost').oninput = $('uplan').oninput = $('uprice').oninput = $('uover').oninput = function(){ syncCostHint(); if (lastUsage) drawUsage(lastUsage); drawMom(); };
  // ---- Month over month: saved finished months, the rest read one at a time
  // A period is named by the month it STARTS in; with a cycle day of 18,
  // '2026-10' is Oct 18 through Nov 17.
  function cycleDay(){ var v = parseInt($('ucycle').value, 10); return v >= 1 && v <= 28 ? v : 1; }
  function monthsBack(k){
    var out = [], t = new Date(), d = new Date(t.getFullYear(), t.getMonth(), 1);
    if (t.getDate() < cycleDay()) d.setMonth(d.getMonth() - 1);
    for (var i = 0; i < k; i++) { out.unshift(d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2)); d.setMonth(d.getMonth() - 1); }
    return out;
  }
  function plabel(m){
    var day = cycleDay(), y = +m.substr(0, 4), mo = +m.substr(5) - 1;
    var MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    if (day == 1) return MON[mo] + ' ' + y;
    var end = new Date(y, mo + 1, day - 1);
    return MON[mo] + ' ' + day + ' to ' + MON[end.getMonth()] + ' ' + end.getDate();
  }
  var momData = null, momCsv = null;
  $('mcsv').onclick = function(){
    if (!momCsv) { $('mstatus').textContent = 'Press Show first.'; return; }
    var c = momCsv, label = $('mmetric').options[$('mmetric').selectedIndex].text, rt = rate();
    var head = ['Group'].concat(c.months.map(plabel)).concat(['Total']);
    var rows = [['Estimated from the PoweredByText message report and TouchPoint. Your PBT invoice is the real number. Showing: '
      + label + (rt && c.k != 'sent' && c.k != 'notTp' ? '. Dollars at $' + rt + ' per credit in the rows marked $.' : '')]];
    rows.push(head);
    c.names.forEach(function(g){
      var vals = c.months.map(function(m){ var r = c.recs[m]; return r && r.groups ? gval(r.groups[g] || {}, c.k) : ''; });
      var tot = vals.reduce(function(a, v){ return a + (+v || 0); }, 0);
      rows.push([g].concat(vals).concat([tot]));
      if (rt && c.k != 'sent' && c.k != 'notTp') rows.push([g + ' $'].concat(vals.map(function(v){ return v === '' || v == null ? '' : (v * rt).toFixed(2); })).concat([(tot * rt).toFixed(2)]));
    });
    c.totals.forEach(function(r){ rows.push(r.map(function(v){ return v == null ? '' : v; })); });
    var name = 'pbt-usage-' + c.k + '-' + c.months[0] + '-to-' + c.months[c.months.length - 1] + '.csv';
    csv(name, rows[0], rows.slice(1));
  };
  $('mmetric').onchange = function(){ drawMom(); };
  function gval(g, k){
    if (!g) return 0;
    if (k == 'waste') return (g.failedCredits || 0) + (g.extra || 0) + (g.dupCredits || 0);
    if (k == 'notTp') return g.tpSent == null ? null : Math.max(0, (g.sent || 0) - (g.tpSent || 0));
    return g[k] || 0;
  }
  // Usage is saved by day. Refresh reads only what is new since the last
  // refresh (plus a few days of overlap); Show adds up the saved days.
  var REFRESH_AFTER_MIN = 60;
  function throughText(t){ return t ? 'Counted through ' + t.substr(5, 11).replace(' ', ' at ') + '.' : 'Not counted yet.'; }
  var MONS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  function md(d){ return MONS[+d.substr(5, 2) - 1] + ' ' + (+d.substr(8, 2)); }
  // Reads PBT a week at a time so the bar can show where it is. Each piece
  // saves its days, so a refresh that stops partway keeps what it read.
  function refreshThen(next){
    $('mstatus').textContent = 'Checking what is new...'; $('mbuild').disabled = $('mrefresh').disabled = true;
    var t0 = Date.now(), rows = 0;
    function fail(msg){ $('mbuild').disabled = $('mrefresh').disabled = false; $('mprog').style.display = 'none'; $('mstatus').textContent = 'Refresh stopped: ' + msg; }
    post({action:'usage_refresh_plan'}, function(p){
      if (!p.success) return fail(p.message);
      var sl = p.slices, i = 0;
      $('mprog').style.display = '';
      (function piece(){
        $('mbar').style.width = Math.round(100 * i / sl.length) + '%';
        if (i >= sl.length) {
          $('mbuild').disabled = $('mrefresh').disabled = false; $('mprog').style.display = 'none';
          $('mstatus').textContent = 'Read ' + n(rows) + ' report rows in ' + Math.round((Date.now() - t0) / 1000) + 's.';
          return next(false);
        }
        $('mprogtxt').textContent = 'Reading ' + md(sl[i][0]) + (sl[i][1] != sl[i][0] ? ' to ' + md(sl[i][1]) : '') + ' from PoweredByText ('
          + (i + 1) + ' of ' + sl.length + '), ' + Math.round((Date.now() - t0) / 1000) + 's so far. Only what is new since the last refresh is read.';
        post({action:'usage_refresh', sfrom:sl[i][0], sto:sl[i][1], firstpiece:i == 0 ? '1' : '0', lastpiece:i == sl.length - 1 ? '1' : '0'}, function(r){
          if (!r.success) return fail(r.message);
          rows += r.rows || 0; i++; piece();
        });
      })();
    });
  }
  // ---- summary cards for the current billing period
  function drawKpi(){
    if (!momData) return;
    var ms = momData.months, cur = ms[ms.length - 1], prev = ms[ms.length - 2], r = momData.recs[cur], pr = momData.recs[prev];
    if (!r || !r.total) { $('ukpi').innerHTML = ''; return; }
    var t = r.total, waste = (t.failedCredits || 0) + (t.extraForSpecial || 0) + (t.repeatCredits || 0);
    var day = cycleDay(), y = +cur.substr(0, 4), mo = +cur.substr(5) - 1;
    var start = new Date(y, mo, day), end = new Date(y, mo + 1, day), now = new Date();
    var frac = Math.min(1, Math.max(0.02, (now - start) / (end - start)));
    var pace = Math.round(t.credits / frac), plan = numv('uplan');
    var card = function(big, small, cls){ return '<div class="card"><b' + (cls ? ' class="' + cls + '"' : '') + '>' + big + '</b><span>' + small + '</span></div>'; };
    var h = card(n(t.credits), 'credits so far, ' + plabel(cur) + (money(t.credits) ? ' (' + money(t.credits) + ')' : ''));
    if (plan) h += card(Math.round(100 * t.credits / plan) + '%', 'of the ' + n(plan) + ' credit plan; on pace for ' + Math.round(100 * pace / plan) + '%', pace > plan ? 'late' : '');
    else h += card(n(pace), 'credits on pace for the whole period');
    h += card(n(waste), 'wasted: ' + n(t.failedCredits || 0) + ' not delivered, ' + n(t.extraForSpecial || 0) + ' emoji, ' + n(t.repeatCredits || 0) + ' sent twice' + (money(waste) ? ' (' + money(waste) + ')' : ''), waste ? 'late' : '');
    h += card(n(t.sent), 'texts sent' + (t.received ? ', ' + n(t.received) + ' received' : ''));
    if (pr && pr.total) h += card(n(pr.total.credits), 'credits last period (' + plabel(prev) + ')');
    $('ukpi').innerHTML = h;
    $('ukpinote').textContent = 'Click any number in the grid below to see the texts and people behind it.';
  }
  $('mrefresh').onclick = function(){ refreshThen(build); };
  $('mbuild').onclick = function(){ build(true); };
  function build(mayRefresh){
    var want = monthsBack(+$('mmonths').value), recount = $('mrecount').checked;
    $('mstatus').textContent = ''; $('mout').innerHTML = '';
    $('mprog').style.display = ''; $('mbar').style.width = '5%';
    $('mprogtxt').textContent = 'Loading the saved billing periods...';
    post({action:'usage_history'}, function(r){
      if (!r.success) { $('mprog').style.display = 'none'; $('mstatus').textContent = r.message; return; }
      var age = r.daysThrough ? (Date.now() - new Date(r.daysThrough.replace(' ', 'T'))) / 60000 : 1e9;
      if (mayRefresh === 'saved' && !r.daysThrough) {
        $('mprog').style.display = 'none'; $('mstatus').textContent = '';
        $('ukpi').innerHTML = '<div class="box" style="flex:1"><b>No usage counted yet.</b> Counting reads the last two months from PoweredByText '
          + '(about a minute for a busy church) and saves it. After that, each refresh reads only the days since the last one. '
          + '<button id="ufirst">Count now</button></div>';
        $('ufirst').onclick = function(){ $('ukpi').innerHTML = ''; refreshThen(build); };
        return;
      }
      if (mayRefresh === true && age > REFRESH_AFTER_MIN) { refreshThen(build); return; }
      var have = r.months || {};
      var day = cycleDay();
      var todo = want.filter(function(m){ return recount || !have[m] || !have[m].complete || (have[m].v || 1) < (r.version || 2)
        || (have[m].cycleDay || 1) != day; });
      var i = 0, t0 = Date.now(), btn = $('mbuild');
      btn.disabled = true;
      $('mprog').style.display = '';
      var tick = setInterval(function(){
        if (i < todo.length) $('mprogtxt').textContent = progText();
      }, 1000);
      function progText(){
        var m = todo[i], el = Math.round((Date.now() - t0) / 1000);
        var recent = m >= monthsBack(3)[0];
        return 'Counting ' + plabel(m) + ' (' + (i + 1) + ' of ' + todo.length + '), ' + el + 's so far. '
          + (recent ? 'Adding up the saved days. '
                    : 'Periods older than the saved days come from TouchPoint and are quick. ')
          + 'Finished periods are saved, so this happens once.';
      }
      (function next(){
        $('mbar').style.width = (todo.length ? Math.round(5 + 95 * i / todo.length) : 100) + '%';
        if (i >= todo.length) {
          clearInterval(tick); btn.disabled = false;
          setTimeout(function(){ $('mprog').style.display = 'none'; }, 300);
          momData = {months: want, recs: have}; drawMom();
          $('mstatus').textContent = (todo.length ? 'Counted ' + todo.length + ' in ' + Math.round((Date.now() - t0) / 1000) + 's. ' : '') + throughText(r.daysThrough);
          return;
        }
        $('mstatus').textContent = '';
        $('mprogtxt').textContent = progText();
        post({action:'usage_month', month:todo[i]}, function(x){
          if (x.success) have[todo[i]] = x.rec;
          else have[todo[i]] = {month: todo[i], error: x.message};
          i++; next();
        });
      })();
    });
  };
  function drawMom(){
    if (!momData) return;
    var ms = momData.months, recs = momData.recs, groups = {};
    var k = $('mmetric').value, isCount = k == 'sent' || k == 'notTp' || k == 'unknown';
    ms.forEach(function(m){ var r = recs[m]; if (r && r.groups) Object.keys(r.groups).forEach(function(g){ groups[g] = (groups[g] || 0) + (gval(r.groups[g], k) || 0); }); });
    var names = Object.keys(groups).filter(function(g){ return groups[g] > 0 || k == 'credits'; }).sort(function(a, b){ return groups[b] - groups[a]; });
    var cell = function(cr, cls, plain, dm){
      var can = dm && cr && DRILLABLE[dm[2]];
      return '<td class="num' + (cls ? ' ' + cls : '') + '"' + (can ? ' data-dm="' + esc(dm[0]) + '" data-dg="' + esc(dm[1]) + '" data-dk="' + dm[2]
        + '" title="Click to see what is behind this number"' : '') + '>' + (cr == null ? '<span class="who">-</span>' : (plain ? n(cr) : crm(cr))) + '</td>'; };
    var head = function(m){
      var r = recs[m], lab = plabel(m), parts = lab.split(' to ');
      return (parts.length == 2 ? esc(parts[0]) + '<div class="who">to ' + esc(parts[1]) + '</div>' : esc(lab))
        + (r && r.complete === false ? '<div class="who">so far</div>' : '')
        + (r && r.source == 'TouchPoint' ? '<span class="tag t-same" title="PBT no longer reports this period. Counted from TouchPoint, which records every text and its delivery result, but not automatic replies.">TP</span>' : '');
    };
    var h = '<div style="overflow-x:auto;max-width:100%"><table class="mom"><thead><tr><th class="g">Group</th>'
      + ms.map(function(m){ return '<th class="num">' + head(m) + '</th>'; }).join('')
      + '<th class="num">Total</th></tr></thead><tbody>';
    names.forEach(function(g){
      h += '<tr><td class="g">' + esc(g) + '</td>' + ms.map(function(m){ var r = recs[m]; return cell(r && r.groups ? gval(r.groups[g] || {sent: 0, tpSent: r.tpTotal ? 0 : null}, k) : null, '', isCount, [m, g, k]); }).join('')
        + cell(groups[g], '', isCount) + '</tr>';
    });
    var sum = function(f){ return ms.map(function(m){ var r = recs[m]; return r && r.total ? f(r.total) : null; }); };
    var csvRows = [];
    var row = function(label, vals, cls, plain, dk){ var tot = vals.reduce(function(a, v){ return a + (v || 0); }, 0);
      csvRows.push([label.replace(/&nbsp;/g, '').trim()].concat(vals).concat([tot]));
      return '<tr><td class="g"><b>' + label + '</b></td>' + vals.map(function(v, i){ return cell(v, cls, plain, dk ? [ms[i], '', dk] : null); }).join('') + cell(tot, cls, plain) + '</tr>'; };
    h += row('All groups, credits', sum(function(t){ return t.credits; }), '', 0, 'credits');
    var planC = numv('uplan'), overR = numv('uover');
    if (planC) {
      var used = sum(function(t){ return t.credits; });
      h += '<tr><td class="g"><b>&nbsp;&nbsp;Plan used</b><div class="who">of ' + n(planC) + ' credits</div></td>' + used.map(function(v, i){
          var r = recs[ms[i]], p = v == null ? null : Math.round(100 * v / planC);
          return '<td class="num' + (p > 100 ? ' late' : '') + '">' + (p == null ? '<span class="who">-</span>' : p + '%'
            + (r && r.complete === false ? '<div class="who">so far</div>' : '')) + '</td>'; }).join('') + '<td></td></tr>';
      csvRows.push(['Plan used (%)'].concat(used.map(function(v){ return v == null ? '' : Math.round(100 * v / planC); })).concat(['']));
      var over = used.map(function(v){ return v == null ? null : Math.max(0, v - planC); });
      h += row('&nbsp;&nbsp;Over plan (credits)', over, 'late', 1);
      if (overR) h += row('&nbsp;&nbsp;Overage cost $', over.map(function(v){ return v == null ? null : Math.round(v * overR * 100) / 100; }), 'late', 1);
    }
    h += row('Waste', sum(function(t){ return (t.failedCredits || 0) + (t.extraForSpecial || 0) + (t.repeatCredits || 0); }), 'late', 0, 'waste');
    h += row('&nbsp;&nbsp;Not delivered', sum(function(t){ return t.failedCredits || 0; }), '', 0, 'failedCredits');
    h += row('&nbsp;&nbsp;Emoji or special characters', sum(function(t){ return t.extraForSpecial || 0; }), '', 0, 'extra');
    h += row('&nbsp;&nbsp;Same person got it twice', sum(function(t){ return t.repeatCredits || 0; }), '', 0, 'dupCredits');
    h += row('Fallback texts (MMS could not go)', sum(function(t){ return t.fallbackCredits || 0; }), '', 0, 'fallbackCredits');
    h += row('No final status (texts)', ms.map(function(m){ var r = recs[m]; return r && r.source == 'TouchPoint' && r.total ? (r.total.unknown || 0) : null; }), '', 1, 'unknown');
    h += row('Refused, out of credits (texts)', ms.map(function(m){ var r = recs[m]; return r && r.source == 'TouchPoint' && r.total ? (r.total.noCredit || 0) : null; }), 'late', 1, 'noCredit');
    h += row('Texts in PBT', ms.map(function(m){ var r = recs[m]; return r && r.source == 'PBT' && r.total ? r.total.sent : null; }), '', 1);
    h += row('&nbsp;&nbsp;Sent from TouchPoint', ms.map(function(m){ var r = recs[m];
      return r && r.tpTotal ? r.tpTotal.sent : (r && r.source == 'TouchPoint' && r.total ? r.total.sent : null); }), '', 1);
    h += row('&nbsp;&nbsp;Not in TouchPoint', ms.map(function(m){ var r = recs[m];
      return r && r.tpTotal && r.total ? Math.max(0, r.total.sent - r.tpTotal.sent) : null; }), 'late', 1);
    momCsv = {months: ms, names: names, recs: recs, k: k, totals: csvRows};
    drawKpi();
    h += '</tbody></table></div>';
    var errs = ms.filter(function(m){ return recs[m] && recs[m].error; });
    if (errs.length) h += '<p class="note">Could not count: ' + errs.map(function(m){ return esc(m + ' (' + recs[m].error + ')'); }).join(', ') + '</p>';
    h += '<p class="note">Each column is one billing period (the start day is under Settings). '
      + 'Estimated credits, counted the same way as below. PBT keeps about two months of report data, so finished periods '
      + 'are saved the first time they are counted (and each morning, if morning updates are on). Periods tagged '
      + '<b>TP</b> come from TouchPoint\\'s record of each recipient and the delivery result it stored. '
      + 'TouchPoint records texts sent from inside PBT too (several times over, as reports arrive, so copies are counted once). '
      + 'It does not record automatic replies. <b>No final status</b> are texts whose result never came back. '
      + 'Waste counts a duplicate only when the same person got the same text again within 3 days; the same wording weeks later is normal reuse. '
      + '<b>Not in TouchPoint</b> is texts PBT sent that TouchPoint has no record of: sent from inside PBT, automatic replies and '
      + 'system messages, only for periods PBT still reports. Fill in your plan above to see dollars and how much of the plan each period uses. '
      + '<b>Click any number</b> to see the texts and people behind it.</p>';
    $('mout').innerHTML = h;
  }
  // ---- Click a number in the grid: the texts and people behind it
  var DRILLABLE = {credits: 'Credits', waste: 'Waste', failedCredits: 'Not delivered', extra: 'Emoji or special characters',
    dupCredits: 'Same person got it twice', fallbackCredits: 'Fallback texts', sent: 'Texts sent', unknown: 'No final status', noCredit: 'Refused, out of credits'};
  var CAUSE = {shared: ['Phone number shared by more than one person', 'TouchPoint sends one text per person on the list. When a parent cell phone is also on the children, that phone gets the text once for each of them, and each one is billed. Fix the cell phone on the child records (the Names tab finds these), or send to parents only.'],
    again: ['Sent again', 'The same text went to the same person a second time within 3 days, from a new send.'],
    retry: ['Tried again after a failure', 'The first one failed and it was sent again. Worth it only if the number has started working.']};
  $('mout').onclick = function(ev){
    var td = ev.target.closest ? ev.target.closest('td[data-dm]') : null;
    if (!td) return;
    Array.prototype.forEach.call(document.querySelectorAll('#mout td.picked'), function(x){ x.classList.remove('picked'); });
    td.classList.add('picked');
    var m = td.getAttribute('data-dm'), g = td.getAttribute('data-dg'), k = td.getAttribute('data-dk');
    var head = '<b>' + esc(DRILLABLE[k]) + '</b>, ' + esc(g || 'All groups') + ', ' + esc(plabel(m));
    $('mdrill').innerHTML = '<div class="box">' + head + '<p class="note">Reading ' + (m >= monthsBack(3)[0] ? 'PBT report rows, this can take a minute...' : 'TouchPoint records...') + '</p></div>';
    $('mdrill').scrollIntoView({behavior: 'smooth', block: 'start'});
    post({action: 'drill', month: m, grp: g, metric: k}, function(r){
      if (!r.success) { $('mdrill').innerHTML = '<div class="box">' + head + '<p class="note">' + esc(r.message) + '</p></div>'; return; }
      $('mdrill').innerHTML = '<div class="box"><button class="ghost" id="mdclose" style="float:right">Close</button>' + head
        + ' <span class="tag t-same">' + (r.source == 'TouchPoint' ? 'from TouchPoint' : 'from PBT') + '</span>' + drillBody(r, k) + '</div>';
      $('mdclose').onclick = function(){ $('mdrill').innerHTML = ''; td.classList.remove('picked'); };
    });
  };
  function person(r){
    return fmt(r.num) + '<div class="who">' + (r.pid ? '<a href="/Person2/' + r.pid + '" target="_blank">' + esc(r.owner) + '</a>' : esc(r.owner || 'not in TouchPoint')) + '</div>';
  }
  function more(r){ return r.more ? '<p class="note">Showing the first ' + r.rows.length + '; ' + n(r.more) + ' more not shown.</p>' : ''; }
  function drillBody(r, k){
    var T = r.total, h = '';
    if (r.kind == 'messages') {
      h += '<p class="note">' + n(r.rows.length + r.more) + ' different texts, ' + n(T.texts) + ' sent, ' + crm(T.credits).replace('<div class="who">', ' (').replace('</div>', ')') + ' credits'
        + (k == 'extra' ? ', ' + n(T.extra) + ' of them extra for emoji or special characters' : '')
        + (k == 'waste' ? ', ' + n(T.waste) + ' wasted' : '') + '. A text sent several times in the period is one row.</p>';
      h += '<table><thead><tr><th>First sent</th><th>Group</th><th>Sent by</th><th class="num">People</th><th class="num">Texts</th>'
        + '<th class="num">Credits each</th><th class="num">Credits</th>'
        + (k == 'waste' ? '<th class="num">Not delivered</th><th class="num">Emoji extra</th><th class="num">Twice</th>' : '')
        + (k == 'extra' ? '<th class="num">Extra</th>' : '') + '<th>Message, and why it costs what it does</th></tr></thead><tbody>'
        + r.rows.map(function(x){
            var chips = (x.special || []).map(function(c){ return '<span class="chip" title="Type ' + esc(c.fix) + ' instead">' + esc(c.ch) + ' ' + esc(c.name)
              + (c.n > 1 ? ' x' + c.n : '') + (c.fix && c.fix.indexOf(' ') < 0 ? ', use ' + esc(c.fix) : '') + '</span>'; }).join('');
            return '<tr><td>' + esc(x.first) + (x.sends > 1 ? '<div class="who">' + x.sends + ' sends</div>' : '') + '</td><td>' + esc(x.group) + '</td><td>' + esc(x.senders)
              + '</td><td class="num">' + n(x.people) + '</td><td class="num">' + n(x.texts) + (x.texts > x.people ? '<div class="who">' + n(x.texts - x.people) + ' repeats</div>' : '')
              + '</td><td class="num">' + x.perPerson + (x.perPerson > x.plain ? '<div class="who">' + x.plain + ' if plain</div>' : '') + '</td><td class="num">' + crm(x.credits) + '</td>'
              + (k == 'waste' ? '<td class="num">' + n(x.failedCredits) + '</td><td class="num">' + n(x.extra) + '</td><td class="num">' + n(x.dupCredits) + '</td>' : '')
              + (k == 'extra' ? '<td class="num late">' + crm(x.extra) + '</td>' : '')
              + '<td class="msg">' + esc(x.text) + '<div class="why">' + esc(x.why) + '</div>' + (k == 'extra' || k == 'waste' ? chips : '') + '</td></tr>';
          }).join('') + '</tbody></table>' + more(r);
      if (k == 'extra') h += '<p class="note">Curly quotes and apostrophes usually come from typing on a phone or pasting from Word or email. '
        + 'One curly apostrophe costs as much as an emoji. Credits each is from the text as written; '
        + 'names and links filled in for each person can change the length a little.</p>';
      return h;
    }
    if (r.kind == 'repeats') {
      h += '<div class="cards">' + r.causes.map(function(c){ return '<div class="card"><b>' + crm(c.credits) + '</b><span>' + esc(CAUSE[c.cause][0])
        + ', ' + n(c.texts) + ' texts</span></div>'; }).join('') + '</div>';
      h += r.causes.map(function(c){ return '<p class="note"><b>' + esc(CAUSE[c.cause][0]) + '.</b> ' + esc(CAUSE[c.cause][1]) + '</p>'; }).join('');
      if (r.sharedSends.length) h += '<p style="margin:10px 0 4px"><b>Sends that reached shared numbers</b></p><table><thead><tr><th>When</th><th>Group</th><th>Sent by</th>'
        + '<th class="num">Numbers hit twice or more</th><th class="num">Extra credits</th><th>Message</th></tr></thead><tbody>'
        + r.sharedSends.map(function(x){ return '<tr><td>' + esc(x.when) + '</td><td>' + esc(x.group) + '</td><td>' + esc(x.sender) + '</td><td class="num">'
          + n(x.numbers) + '</td><td class="num">' + crm(x.credits) + '</td><td class="msg">' + esc(x.text.substr(0, 140)) + '</td></tr>'; }).join('') + '</tbody></table>';
      h += '<p style="margin:10px 0 4px"><b>Each repeat</b></p><table><thead><tr><th>Number</th><th>Group</th><th>First</th><th>Again</th><th>Why</th><th class="num">Credits</th><th>Message</th></tr></thead><tbody>'
        + r.rows.map(function(x){ return '<tr><td>' + person(x) + '</td><td>' + esc(x.group) + '</td><td>' + esc(x.first) + '</td><td>' + esc(x.again)
          + '</td><td>' + esc(x.why) + '</td><td class="num">' + x.credits + '</td><td class="msg">' + esc(x.text.substr(0, 120)) + '</td></tr>'; }).join('')
        + '</tbody></table>' + more(r);
      return h;
    }
    h += '<p class="note">' + n(T.texts) + ' texts to ' + n(T.numbers) + ' numbers, ' + n(T.credits) + ' credits.</p>';
    h += '<table style="margin-bottom:12px"><thead><tr><th>Status</th><th class="num">Texts</th><th class="num">Numbers</th><th class="num">Credits</th></tr></thead><tbody>'
      + r.statuses.map(function(s){ return '<tr><td>' + esc(s.status) + '</td><td class="num">' + n(s.texts) + '</td><td class="num">' + n(s.numbers) + '</td><td class="num">' + crm(s.credits) + '</td></tr>'; }).join('')
      + '</tbody></table>';
    h += '<table><thead><tr><th>Number</th><th class="num">Times</th><th>When</th><th>Status</th><th>Group</th><th>Sent by</th><th>Message</th></tr></thead><tbody>'
      + r.rows.map(function(x){ return '<tr><td>' + person(x) + '</td><td class="num">' + x.timesThisNumber + '</td><td>' + esc(x.when)
        + (x.late ? '<div class="who">late result</div>' : '') + '</td><td class="k-failed">' + esc(x.status) + '</td><td>' + esc(x.group) + '</td><td>' + esc(x.sender)
        + '</td><td class="msg">' + esc(x.text.substr(0, 120)) + '</td></tr>'; }).join('') + '</tbody></table>' + more(r)
      + '<p class="note">Times is how often this number shows up in this list for the period. A number that keeps failing is worth fixing or removing in TouchPoint.</p>';
    return h;
  }
  $('uload').onclick = function(){
    $('ucstatus').textContent = ''; $('uout').innerHTML = ''; $('unotes').textContent = '';
    var a = $('ufrom').value, b = $('uto').value;
    var done = waitBar('uc', 'Reading ' + rangeText(a, b) + ' from PoweredByText', rangeDays(a, b), 'uload');
    post({action:'usage', gid:$('ugroup').value, from:a, to:b}, function(r){
      var took = done();
      if (!r.success) { $('ucstatus').textContent = r.message; return; }
      $('ucstatus').textContent = n(r.rowCount) + ' report rows, in ' + took + 's.';
      $('unotes').textContent = (r.notes || []).join(' ');
      lastUsage = r;
      drawUsage(r);
    });
  };
  function drawUsage(r){
      var t = r.total;
      var tbl = function(head, rows){ return '<table style="margin-bottom:14px"><thead><tr>' + head.map(function(h){
          return '<th' + (h.charAt(0) == '#' ? ' class="num">' + h.substr(1) : '>' + h) + '</th>'; }).join('') + '</tr></thead><tbody>' + rows.join('') + '</tbody></table>'; };
      var td = function(v, num, cls){ return '<td class="' + (num ? 'num ' : '') + (cls || '') + '">' + v + '</td>'; };
      var waste = t.failedCredits + t.extraForSpecial + (t.repeatCredits || 0);
      var h = '<div class="box"><b>Waste</b> <span class="who">credits spent that did not need to be, in this range</span>'
        + '<div class="cards">'
        + '<div class="card"><b class="late">' + crm(waste) + '</b><span>total, estimated</span></div>'
        + '<div class="card"><b>' + crm(t.failedCredits) + '</b><span>on ' + n(t.failed) + ' texts not delivered</span></div>'
        + '<div class="card"><b>' + crm(t.repeatFailCredits || 0) + '</b><span>of that, on numbers failing 3+ times (' + n(t.repeatFailNumbers || 0) + ')</span></div>'
        + '<div class="card"><b>' + crm(t.extraForSpecial) + '</b><span>extra from emoji or special characters</span></div>'
        + '<div class="card"><b>' + crm(t.repeatCredits || 0) + '</b><span>same text sent again (some on purpose)</span></div>'
        + '</div>' + (rate() ? '' : '<div class="note">Fill in your plan under Settings to see dollars.</div>') + '</div>'
        + '<div class="cards">'
        + '<div class="card"><b>' + crm(t.credits) + '</b><span>estimated credits</span></div>'
        + '<div class="card"><b>' + n(t.sent) + '</b><span>texts sent</span></div>'
        + '<div class="card"><b>' + n(t.mms) + '</b><span>picture messages (3 credits each)</span></div>'
        + '<div class="card"><b>' + n(t.fallbackCredits) + '</b><span>fallback texts when a picture could not go</span></div>'
        + '<div class="card"><b>' + n(t.received) + '</b><span>texts received (not priced)</span></div></div>';
      h += '<p class="note">Credits follow PBT\\'s help pages: 1 per text segment (160 characters, or 70 when the text has emoji or other '
        + 'special characters), a text over 3 segments goes as a text-only MMS, and any MMS is 3 credits. Estimates; your PBT invoice is the real number. '
        + 'Undelivered texts are counted because carriers usually still charge for them.</p>';
      if ((r.costly || []).length) {
        h += '<h4>Where credits go: the costliest texts</h4>' + tbl(['First out', 'Group', '#Sends', '#Recipients', '#Credits', '#Each', 'Why', 'Message'],
          r.costly.map(function(e){ return '<tr>' + td(esc(e.first)) + td(esc(e.group)) + td(n(e.sends), 1) + td(n(e.recipients), 1)
            + td(crm(e.credits), 1) + td(n(e.perPerson), 1) + td(esc(e.why) + (e.extra ? '<div class="late">' + n(e.extra) + ' extra from emoji or special characters</div>' : ''))
            + td(esc(e.message), 0, 'msg') + '</tr>'; }));
      }
      if ((r.repeats || []).length) {
        h += '<h4>Same text sent more than once</h4><p class="note">The same message to the same group at separate times. Some are on purpose (a reminder); some are a send repeated by mistake.</p>'
          + tbl(['First out', 'Group', '#Times sent', '#Recipients', '#Credits', 'Message'],
          r.repeats.map(function(e){ return '<tr>' + td(esc(e.first)) + td(esc(e.group)) + td(n(e.sends), 1, 'late') + td(n(e.recipients), 1)
            + td(crm(e.credits), 1) + td(esc(e.message), 0, 'msg') + '</tr>'; }));
      }
      if ((r.failing || []).length) {
        h += '<h4>Numbers that keep failing</h4><p class="note">Three or more undelivered texts in this range. Usually a landline, a disconnected '
          + 'number, or a carrier block. Every text to them still uses credits; fix or remove the number in TouchPoint.</p>'
          + tbl(['Number', 'Whose number', '#Failed', '#Credits', 'Groups', 'Last'],
          r.failing.map(function(f){ return '<tr>' + td(fmt(f.num)) + td(who([f.num, f.name || '', '', '', '', f.pid, f.why])) + td(n(f.failed), 1, 'k-failed')
            + td(crm(f.credits), 1) + td(esc(f.groups)) + td(esc((f.last || '').substr(5, 11))) + '</tr>'; }));
      }
      if ((r.senders || []).length) {
        h += '<h4>Who sends the most</h4><p class="note">From TouchPoint\\'s record of who pressed Send. PBT does not report a sender, '
          + 'so texts sent from inside PBT and automatic replies are one line at the bottom.</p>'
          + tbl(['Sender', '#Texts', '#Recipients', '#Est. credits', 'Mostly to'],
          r.senders.map(function(x){ return '<tr>' + td(esc(x.sender)) + td(x.texts == null ? '' : n(x.texts), 1) + td(n(x.sent), 1)
            + td(x.credits == null ? '' : crm(x.credits), 1) + td(esc(x.groups)) + '</tr>'; }));
      }
      h += '<h4>By group, most first</h4>' + tbl(['Group', '#Credits', '#Sent', '#People reached', '#Pictures', '#Received', '#Not delivered'],
        r.groups.map(function(g){ return '<tr>' + td(esc(g.group)) + td(crm(g.credits), 1) + td(n(g.sent), 1) + td(n(g.people), 1)
          + td(n(g.mms), 1) + td(n(g.received), 1) + td(n(g.failed), 1) + '</tr>'; }));
      if ((r.months || []).length > 1) {
        h += '<h4>By month</h4>' + tbl(['Month', '#Credits', '#Sent', '#Pictures', '#Received'],
          r.months.map(function(m){ return '<tr>' + td(esc(m.month)) + td(crm(m.credits), 1) + td(n(m.sent), 1) + td(n(m.mms), 1) + td(n(m.received), 1) + '</tr>'; }));
      }
      h += '<h4>By kind of text</h4>' + tbl(['Kind', '#Credits', '#Sent'],
        r.types.map(function(x){ return '<tr>' + td(esc(x.type)) + td(crm(x.credits), 1) + td(n(x.sent), 1) + '</tr>'; }));
      $('uout').innerHTML = h;
  }

  $('run').onclick = function(){
    var ids = boxes().filter(function(b){ return b.checked; }).map(function(b){
      return {id:b.value, name:b.parentNode.textContent.trim()}; });
    if (!ids.length || busy) return;
    merged = {}; filter = null; sel = {};
    var btn = this, log = $('log'), done = 0, errs = [], t0 = Date.now();
    busy = true; btn.disabled = true; $('csv').disabled = true;
    log.textContent = ''; $('nprog').style.display = '';
    (function next(){
      $('nbar').style.width = Math.round(100 * done / ids.length) + '%';
      if (done == ids.length) {
        busy = false; btn.disabled = false; $('csv').disabled = false;
        setTimeout(function(){ $('nprog').style.display = 'none'; }, 300);
        log.innerHTML = 'Checked ' + ids.length + ' group(s) in ' + Math.round((Date.now() - t0) / 1000) + 's.' + (errs.length ? ' <span class="kid">Problems: ' + esc(errs.join('; ')) + '</span>' : '');
        render(); return;
      }
      var g = ids[done];
      $('nprogtxt').textContent = 'Reading ' + g.name + ' from PoweredByText (' + (done + 1) + ' of ' + ids.length + '), '
        + Math.round((Date.now() - t0) / 1000) + 's so far. Results fill in below as each group finishes.';
      post({action:'preview_group', gid:g.id}, function(r){
        if (r.success) {
          r.rows.forEach(function(row){
            // Contacts are per brand (orgunit), so one number = one name per brand.
            var key = r.orgunitId + '|' + row.num;
            if (!merged[key]) { row.groups = []; row.key = key; merged[key] = row; }
            merged[key].groups.push(r.group);
          });
        } else errs.push(g.name + ': ' + r.message);
        done++; render(); next();
      });
    })();
  };

  function list(){
    var q = ($('nq').value || '').toLowerCase();
    return Object.keys(merged).map(function(k){ return merged[k]; }).filter(function(r){
      if (filter == 'multi') { if (!multi(r)) return false; }
      else if (filter && r.action != filter) return false;
      if (!q) return true;
      return (r.num + ' ' + r.pbt + ' ' + r.first + ' ' + r.last + ' ' +
              r.cands.map(function(c){ return c.name; }).join(' ')).toLowerCase().indexOf(q) >= 0;
    }).sort(function(a, b){ return ORDER.indexOf(a.action) - ORDER.indexOf(b.action) || (a.last + a.first).localeCompare(b.last + b.first); });
  }
  function selCount(){ return Object.keys(sel).filter(function(k){ return sel[k] && merged[k] && changeable(merged[k]); }).length; }
  function buttons(){
    var all = Object.keys(merged).map(function(k){ return merged[k]; }).filter(changeable).length, n = selCount();
    $('applysel').textContent = 'Change selected (' + n + ')';
    $('applysel').disabled = busy || !n;
    $('applyall').textContent = 'Change all Will set and Will fix (' + all + ')';
    $('applyall').disabled = busy || !all;
  }
  function render(){
    var all = Object.keys(merged).map(function(k){ return merged[k]; }), counts = {};
    all.forEach(function(r){ counts[r.action] = (counts[r.action] || 0) + 1; });
    $('cards').innerHTML =
      '<div class="card' + (filter ? '' : ' on') + '" data-a=""><b>' + all.length + '</b><span>Numbers</span></div>' +
      ORDER.map(function(a){ return '<div class="card' + (filter == a ? ' on' : '') + '" data-a="' + a + '"><b>' + (counts[a] || 0) +
        '</b><span class="tag t-' + a + '">' + LABEL[a] + '</span><br><span>' + HELP[a] + '</span></div>'; }).join('') +
      (WRITES ? '' : '<div class="card' + (filter == 'multi' ? ' on' : '') + '" data-a="multi"><b>' + all.filter(multi).length +
        '</b><span class="tag t-multi">' + LABEL.multi + '</span><br><span>' + HELP.multi + '</span></div>');
    Array.prototype.forEach.call(document.querySelectorAll('#cards .card'), function(c){
      c.onclick = function(){ filter = this.getAttribute('data-a') || null; render(); }; });
    var rows = list(), shown = rows.slice(0, 2000);
    $('rows').innerHTML = shown.map(function(r){
      var who = r.cands.map(function(c){
        return '<a href="/Person2/' + c.pid + '" target="_blank"' + (c.adult ? '' : ' class="kid"') + '>' + esc(c.name) + '</a>' +
               (c.age != null ? ' (' + c.age + ')' : '') + (c.pid == r.pid ? ' &#10003;' : ''); }).join('<br>');
      var box = changeable(r) ? '<input type="checkbox" class="pick" data-k="' + esc(r.key) + '"' + (sel[r.key] ? ' checked' : '') + '>' : '';
      var res = r.result ? '<div class="r-' + r.result + '">' + RESULT[r.result] + '</div>' +
        (r.msg ? '<div class="who">' + esc(r.msg) + '</div>' : '') : '';
      return '<tr><td>' + box + '</td><td>' + fmt(r.num) + (r.status && r.status.toLowerCase() != 'active' ? '<div class="who">' + esc(r.status) + '</div>' : '') +
        '</td><td class="who">' + esc(r.groups.join(', ')) + '</td><td>' + esc(r.pbt) +
        '</td><td><span class="tag t-' + r.action + '">' + LABEL[r.action] + '</span>' + res +
        (!WRITES && multi(r) ? '<div><span class="tag t-multi">' + r.cands.length + ' records</span></div>' : '') + '</td><td>' +
        esc((r.first + ' ' + r.last).trim()) + '</td><td class="who">' + esc(r.why) + '</td><td>' + who + '</td></tr>';
    }).join('') + (rows.length > shown.length ? '<tr><td colspan="8"><i>Showing 2,000 of ' + rows.length +
      '. Filter, or download the CSV for all of them.</i></td></tr>' : '');
    Array.prototype.forEach.call(document.querySelectorAll('#rows .pick'), function(b){
      b.onchange = function(){ sel[this.getAttribute('data-k')] = this.checked; buttons(); }; });
    $('selall').checked = false;
    buttons();
  }
  $('nq').oninput = render;
  $('selall').onchange = function(){
    var on = this.checked;
    list().slice(0, 2000).forEach(function(r){ if (changeable(r)) sel[r.key] = on; });
    var keep = this.checked; render(); $('selall').checked = keep;
  };

  // Sends the chosen rows one group at a time, 100 numbers per request. The
  // server reads PoweredByText again and re-decides before writing, so a row
  // that became Keep since the preview is skipped, not overwritten.
  function apply(rows){
    if (!rows.length || busy) return;
    var byGroup = {};
    rows.forEach(function(r){ (byGroup[r.gid] = byGroup[r.gid] || []).push(r); });
    var jobs = [];
    Object.keys(byGroup).forEach(function(gid){
      var rs = byGroup[gid];
      for (var i = 0; i < rs.length; i += 100) jobs.push({gid:gid, rows:rs.slice(i, i + 100)});
    });
    busy = true; buttons(); $('run').disabled = true;
    var done = 0, tally = {changed:0, unconfirmed:0, error:0, skipped:0}, st = $('applystatus');
    (function next(){
      if (!jobs.length) {
        busy = false; $('run').disabled = false; render();
        st.innerHTML = 'Done: <span class="r-changed">' + tally.changed + ' changed</span>' +
          (tally.unconfirmed ? ', <span class="r-unconfirmed">' + tally.unconfirmed + ' not confirmed</span>' : '') +
          (tally.error ? ', <span class="r-error">' + tally.error + ' errors</span>' : '') +
          (tally.skipped ? ', ' + tally.skipped + ' skipped' : '') + '. Every change is in the Change log.';
        return;
      }
      var job = jobs.shift(), byNum = {};
      job.rows.forEach(function(r){ byNum[r.num] = r; });
      st.textContent = 'Changing ' + (done + 1) + ' to ' + (done + job.rows.length) + ' of ' + rows.length + '...';
      post({action:'apply', gid:job.gid, nums:job.rows.map(function(r){ return r.num; }).join(',')}, function(res){
        if (res.success) {
          res.results.forEach(function(x){
            var r = byNum[x.num]; if (!r) return;
            r.result = x.result; r.msg = x.msg; tally[x.result] = (tally[x.result] || 0) + 1;
            if (x.result == 'changed') { r.pbt = x.after; sel[r.key] = false; }
          });
        } else {
          job.rows.forEach(function(r){ r.result = 'error'; r.msg = res.message; });
          tally.error += job.rows.length;
        }
        done += job.rows.length; render(); next();
      });
    })();
  }
  $('applysel').onclick = function(){
    var rows = Object.keys(sel).filter(function(k){ return sel[k] && merged[k] && changeable(merged[k]); })
      .map(function(k){ return merged[k]; });
    if (!rows.length) return;
    if (!confirm('Change ' + rows.length + ' name(s) in PoweredByText?')) return;
    apply(rows);
  };
  $('applyall').onclick = function(){
    var rows = Object.keys(merged).map(function(k){ return merged[k]; }).filter(changeable);
    if (!rows.length) return;
    if (!confirm('Change all ' + rows.length + ' Will set and Will fix names in PoweredByText?\\n\\nTry a few with Change selected first if you have not yet.')) return;
    apply(rows);
  };
  $('csv').onclick = function(){
    csv('PBT-name-sync-preview.csv',
      ['Number','Groups','Name in PBT now','Sync would','New first name','New last name','PeopleId','Why','Result','People in TouchPoint with this number'],
      list().map(function(r){
        return [r.num, r.groups.join('; '), r.pbt, LABEL[r.action], r.first, r.last, r.pid || '', r.why,
                r.result ? RESULT[r.result] + (r.msg ? ': ' + r.msg : '') : '',
                r.cands.map(function(c){ return c.name + (c.age != null ? ' (' + c.age + ')' : '') + ' #' + c.pid; }).join('; ')];
      }));
  };

  $('saveauto').onclick = function(){
    var on = $('auto').checked;
    var groups = boxes().filter(function(b){ return b.checked; }).map(function(b){ return b.value; });
    if (on && !groups.length) { $('autostatus').textContent = 'Check at least one group first.'; return; }
    var parts = window.location.pathname.split('/'), name = '';
    for (var i = 0; i < parts.length; i++)
      if (parts[i] == 'PyScript' || parts[i] == 'PyScriptForm') { name = decodeURIComponent((parts[i + 1] || '').split('?')[0]); break; }
    if (on && !confirm('Add this script to the TouchPoint morning batch?\\n\\nIt adds a small block to the MorningBatch script. Anything else in that script is left alone.')) return;
    $('autostatus').textContent = 'Saving...';
    post({action:'save_settings', auto:on ? '1' : '0', groups:groups.join(','), script_name:name}, function(r){
      $('autostatus').textContent = r.success
        ? (on ? 'Saved. Runs every morning for ' + groups.length + ' group(s).' : 'Saved. Morning updates are off.') +
          (r.installed ? ' In the morning batch.' : ' Not in the morning batch.')
        : 'Not saved: ' + r.message;
    });
  };

  function loadHistory(){
    $('hstatus').textContent = 'Loading...';
    post({action:'log'}, function(r){
      if (!r.success) { $('hstatus').textContent = r.message; return; }
      hist = r.log; renderHistory();
    });
  }
  function hlist(){
    var q = ($('hq').value || '').toLowerCase();
    return hist.filter(function(e){
      return !q || [e.when, e.by, e.num, e.group, e.before, e.after, e.result, e.msg].join(' ').toLowerCase().indexOf(q) >= 0; });
  }
  function renderHistory(){
    var rows = hlist(), shown = rows.slice(0, 1000);
    $('hstatus').textContent = hist.length ? rows.length + ' of ' + hist.length + ' entries, newest first.' : 'No changes yet.';
    $('hrows').innerHTML = shown.map(function(e){
      return '<tr><td>' + esc(e.when) + '</td><td>' + esc(e.by) + '</td><td>' + fmt(e.num) +
        '</td><td class="who">' + esc(e.group) + '</td><td>' + esc(e.before) + '</td><td>' +
        (e.pid ? '<a href="/Person2/' + e.pid + '" target="_blank">' + esc(e.after) + '</a>' : esc(e.after)) +
        '</td><td class="r-' + esc(e.result) + '">' + (RESULT[e.result] || esc(e.result)) + '</td><td class="who">' + esc(e.msg) + '</td></tr>';
    }).join('');
  }
  $('hq').oninput = renderHistory;
  $('hcsv').onclick = function(){
    csv('PBT-name-sync-log.csv', ['When','By','Where','Number','Group','Before','After','PeopleId','Result','Note'],
      hlist().map(function(e){ return [e.when, e.by, e.source, e.num, e.group, e.before, e.after, e.pid || '', RESULT[e.result] || e.result, e.msg]; }));
  };
  buttons();
})();
</script>"""

    return css + body + js


# Opened at /PyScript/<name>? Go to /PyScriptForm/<name>, keeping whatever
# name the script was saved under, so a renamed copy still works.
REDIRECT_JS = """<script>
if (window.location.pathname.indexOf('/PyScript/') > -1)
  window.location.replace(window.location.pathname.replace('/PyScript/', '/PyScriptForm/') + window.location.search + window.location.hash);
</script>"""

# Admin tools actions. Each one checks the Admin role itself (admin_post).
ADMIN_ACTIONS = set(['preview_group', 'apply', 'usage_history', 'usage_month', 'usage_refresh', 'usage_refresh_plan',
                     'save_cost', 'drill', 'pbt_check', 'pbt_check_links', 'apply_update',
                     'trace', 'usage', 'log', 'save_settings'])

# Admin menu: (group, [(page key, label, what it shows)]). 'people' is the
# home screen's subscriber compare; the rest are the admin section's tabs.
ADMIN_MENU = [
    ('People and numbers', [('names', 'Names', 'sync'), ('subs', 'Subscribers', 'people')]),
    ('Sending', [('delivery', 'Delivery', 'delivery'), ('usage', 'Usage and cost', 'usage')]),
    ('Records', [('log', 'Change log', 'history'), ('check', 'Connection check', 'check'), ('settings', 'Settings', 'settings')]),
]

SHELL_CSS = """
<style>
.page-header{display:none!important}
#txapp{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#1f2937}
#txapp .apphead{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:0 0 10px}
#txapp .apphead h2{margin:0;font-size:20px;font-weight:600}
#txapp .apphead .sub{font-size:13px;color:#6b7280}
#txapp .apphead .sp{flex:1}
#txapp .apphead button{padding:6px 12px;border-radius:6px;border:1px solid #1e3a5f;background:#fff;color:#1e3a5f;cursor:pointer;font-size:13px}
#txapp .body{display:flex;gap:18px;align-items:flex-start}
#txapp .main{flex:1;min-width:0;max-width:100%}
#txapp,#txapp .body{max-width:100%;min-width:0}
#txapp nav.adm{width:210px;flex:none;position:sticky;top:10px;background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:8px;display:none}
#txapp.admin nav.adm{display:block}
#txapp nav.adm .grp{font-size:12px;color:#6b7280;padding:10px 10px 3px;margin-top:6px;border-top:1px solid #eef0f3}
#txapp nav.adm .grp:first-child{margin-top:0;border-top:0;padding-top:4px}
#txapp nav.adm a{display:block;padding:6px 10px;border-radius:6px;color:#1f2937;text-decoration:none;font-size:14px}
#txapp nav.adm a:hover{background:#f3f4f6}
#txapp nav.adm a.on{background:#e6edf6;color:#1e3a5f;font-weight:600}
#txapp .txupdate{margin:0 0 10px;padding:10px 12px;border-radius:8px;background:#e6edf6;color:#1e3a5f;font-size:13px}
#txapp .txupdate button{margin-left:6px;padding:4px 12px;border-radius:6px;border:1px solid #1e3a5f;background:#1e3a5f;color:#fff;cursor:pointer;font-size:13px}
#txapp .txupdate button.later{background:#fff;color:#1e3a5f}
#txapp .txcredit{margin:14px 0 4px;text-align:center;font-size:12px;color:#9ca3af}
#txapp .txcredit a{color:#6b7280}
#txapp .txmsg{padding:20px;background:#fff;border:1px solid #e5e7eb;border-radius:10px}
#txapp #tx #tabs,#txapp #pbt .tabs{display:none}
#txapp.home #pbt{display:none}
#txapp.admin.on-pbt #tx{display:none}
#txapp.admin.on-tx #pbt{display:none}
@media (max-width:820px){
  #txapp .apphead .sub{display:none}
  #txapp .apphead h2{font-size:18px}
}
#txapp .menubtn{display:none;align-items:center;gap:8px;padding:6px 12px;border-radius:6px;border:1px solid #d1d5db;background:#fff;color:#1f2937;cursor:pointer;font-size:14px}
#txapp .menubtn .bars{font-size:16px;line-height:1}
/* Narrower screens (most laptops once TouchPoint's margins are counted):
   the menu folds into a button showing the current page. */
/* Phones: a wide admin table scrolls sideways in its own box, so it never
   makes the whole page wider than the screen. */
@media (max-width:820px){
  #txapp #pbt{max-width:100%;min-width:0}
  #txapp #pbt table{display:block;overflow-x:auto;max-width:100%;-webkit-overflow-scrolling:touch}
  #txapp #pbt th{white-space:nowrap}
  #txapp #pbt td{min-width:5.5em}
  #txapp #pbt td.msg{min-width:200px}
}
@media (max-width:1180px){
  #txapp .body{flex-direction:column;gap:10px;align-items:stretch}
  #txapp .main{width:100%}
  #txapp.admin .menubtn{display:inline-flex}
  #txapp.admin nav.adm{display:none;width:100%;max-width:320px;position:static;box-shadow:0 6px 18px rgba(17,24,39,.12)}
  #txapp.admin.menuopen nav.adm{display:block}
}
</style>"""

# Raw string: no escape handling, so the JavaScript arrives exactly as written.
SHELL_JS = r"""
<script>
(function(){
  var app = document.getElementById('txapp');
  if (!app) return;
  var hasInbox = app.getAttribute('data-inbox') == '1', isAdmin = app.getAttribute('data-admin') == '1';
  var MENU = __MENU__;
  function $(id){ return document.getElementById(id); }
  function clickTab(sel){ var a = document.querySelector(sel); if (a) a.click(); }
  function show(page){
    var admin = isAdmin && (page != 'home' || !hasInbox);
    if (admin && page == 'home') page = 'names';
    var item = null;
    MENU.forEach(function(g){ g[1].forEach(function(it){ if (it[0] == page) item = it; }); });
    if (admin && !item) { page = 'names'; item = MENU[0][1][0]; }
    app.className = admin ? 'admin ' + (item[2] == 'people' ? 'on-tx' : 'on-pbt') : 'home';
    if (admin) $('txmenulbl').textContent = item[1];
    $('txmenubtn').setAttribute('aria-expanded', 'false');
    Array.prototype.forEach.call(document.querySelectorAll('#txapp nav.adm a'), function(a){
      a.className = a.getAttribute('data-p') == page ? 'on' : ''; });
    $('txtoadmin').style.display = isAdmin && !admin ? '' : 'none';
    $('txtohome').style.display = admin && hasInbox ? '' : 'none';
    $('txtitle').textContent = admin ? 'Texting admin tools' : 'Texting';
    $('txsub').textContent = admin ? 'Only people with the Admin role see these.' : 'Conversations from your PoweredByText groups';
    if (!admin) clickTab('#tabs a[data-t="inbox"]');
    else if (item[2] == 'people') clickTab('#tabs a[data-t="people"]');
    else clickTab('#pbt .tabs a[data-tab="' + item[2] + '"]');
    var want = admin ? '#admin-' + page : '';
    if ((location.hash || '') != want) history.replaceState(null, '', location.pathname + location.search + want);
  }
  if (isAdmin) {
    $('txnav').innerHTML = MENU.map(function(g){
      return '<div class="grp">' + g[0] + '</div>' + g[1].filter(function(it){ return hasInbox || it[2] != 'people'; }).map(function(it){
        return '<a href="#admin-' + it[0] + '" data-p="' + it[0] + '">' + it[1] + '</a>'; }).join('');
    }).join('');
    Array.prototype.forEach.call(document.querySelectorAll('#txapp nav.adm a'), function(a){
      a.onclick = function(e){ e.preventDefault(); show(a.getAttribute('data-p')); }; });
  }
  $('txmenubtn').onclick = function(){
    var open = app.className.indexOf('menuopen') < 0;
    app.className = app.className.replace(/\s*menuopen/, '') + (open ? ' menuopen' : '');
    this.setAttribute('aria-expanded', open ? 'true' : 'false');
  };
  // ---- updates: admins are told when a newer TPxi Texting is out
  var UPD = __UPD__;
  function newer(pub, mine){
    var a = String(pub || '').split('.'), b = String(mine || '').split('.');
    for (var i = 0; i < Math.max(a.length, b.length); i++){
      var x = parseInt(a[i], 10) || 0, y = parseInt(b[i], 10) || 0;
      if (x !== y) return x > y;
    }
    return false;
  }
  function scriptName(){ var m = location.pathname.match(/\/PyScript(?:Form)?\/([^\/?#]+)/); return m ? decodeURIComponent(m[1]) : UPD.id; }
  if (isAdmin) {
    var skip = ''; try { skip = sessionStorage.getItem('txupdate.later') || ''; } catch(e) {}
    var vx = new XMLHttpRequest();
    vx.open('GET', UPD.api + '/script-versions', true); vx.timeout = 5000;
    vx.onload = function(){
      try {
        var latest = JSON.parse(vx.responseText)[UPD.id];
        if (!latest || !newer(latest, UPD.version) || skip == latest) return;
        var b = $('txupdate');
        b.innerHTML = '<b>A newer TPxi Texting is out.</b> You have ' + UPD.version + ', the newest is ' + latest
          + '. Your settings, conversations and saved counts stay as they are. '
          + '<button type="button" id="txupdgo">Update now</button> <button type="button" class="later" id="txupdlater">Later</button>'
          + ' <span id="txupdst"></span>';
        b.style.display = '';
        $('txupdlater').onclick = function(){ try { sessionStorage.setItem('txupdate.later', latest); } catch(e) {} b.style.display = 'none'; };
        $('txupdgo').onclick = function(){
          if (!confirm('Update TPxi Texting from ' + UPD.version + ' to ' + latest + '?')) return;
          var go = this; go.disabled = true; $('txupdst').textContent = 'Updating...';
          var x = new XMLHttpRequest();
          x.open('POST', location.pathname.replace('/PyScript/', '/PyScriptForm/'), true);
          x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
          x.onload = function(){
            var r; try { r = JSON.parse(x.responseText); } catch(e) { r = {success: false, message: 'Unexpected reply.'}; }
            $('txupdst').textContent = r.message || '';
            if (r.success) setTimeout(function(){ location.reload(); }, 800); else go.disabled = false;
          };
          x.onerror = function(){ go.disabled = false; $('txupdst').textContent = 'Network error.'; };
          x.send('action=apply_update&script_name=' + encodeURIComponent(scriptName()));
        };
      } catch(e) {}
    };
    vx.send();
  }
  $('txtoadmin').onclick = function(){ show('names'); };
  $('txtohome').onclick = function(){ show('home'); };
  var h = (location.hash || '').replace('#admin-', '');
  show(location.hash.indexOf('#admin-') == 0 ? h : 'home');
})();
</script>"""


# The one line at the bottom of every page.
CREDIT_HTML = ('<div class="txcredit">TPxi Texting %s &middot; Free from TPxi. '
               '<a href="https://tpxigo.com" target="_blank" rel="noopener">TPxi Go</a> puts texting on your phone.</div>'
               % TEXTING_VERSION)


def app_page():
    is_admin = model.UserIsInRole('Admin')
    home_html, has_inbox = inbox_page()
    admin_html = admin_page() if is_admin else ''
    if not has_inbox and not is_admin:
        return SHELL_CSS + '<div id="txapp" class="home">' + home_html + CREDIT_HTML + '</div>'
    head = ('<div class="txupdate" id="txupdate" style="display:none"></div>'
            '<div class="apphead"><h2 id="txtitle">Texting</h2><span class="sub" id="txsub"></span><span class="sp"></span>'
            '<button type="button" id="txtoadmin" style="display:none">Admin tools</button>'
            '<button type="button" id="txtohome" style="display:none">Back to conversations</button></div>')
    menu = [[g, [list(it) for it in items]] for g, items in ADMIN_MENU]
    return (SHELL_CSS
            + '<div id="txapp" class="home" data-inbox="%s" data-admin="%s">' % ('1' if has_inbox else '0', '1' if is_admin else '0')
            + head + '<div class="body"><button type="button" class="menubtn" id="txmenubtn" aria-expanded="false" aria-controls="txnav">'
            '<span class="bars" aria-hidden="true">&#9776;</span> <span id="txmenulbl">Menu</span></button>'
            '<nav class="adm" id="txnav" aria-label="Admin tools"></nav><div class="main">'
            + (home_html if has_inbox else '') + admin_html
            + '</div></div>' + CREDIT_HTML + '</div>'
            + SHELL_JS.replace('__MENU__', safe_json(menu)).replace('__UPD__', safe_json(
                {'version': TEXTING_VERSION, 'id': DC_SCRIPT_ID, 'api': DC_API_BASE})))


# ---------------------------------------------------------------- entry
if model.FromMorningBatch:
    if _param('run_pbt_name_sync') == 'true':
        save_last_month()
        morning_run()

elif model.HttpMethod == "post":
    if _param('action') in ADMIN_ACTIONS:
        admin_post()
    else:
        inbox_post()

elif _param('view') == 'pbttest':
    model.Header = ""
    model.Form = REDIRECT_JS + (PBT_TEST_PAGE if model.UserIsInRole('Admin') else
                                '<div style="font-family:sans-serif;padding:20px">Admin only.</div>')

else:
    model.Form = REDIRECT_JS + app_page()
