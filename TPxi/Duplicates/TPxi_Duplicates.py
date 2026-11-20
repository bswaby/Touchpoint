#Roles=Admin,Manager2
# -*- coding: utf-8 -*-
#####################################################################
# TPxi_Duplicates
#####################################################################
# Written By: Ben Swaby
# Email: bswaby@fbchtn.org
# GitHub: https://github.com/bswaby/Touchpoint  (40+ free tools)
# ----------------------------------------------------------------
# These tools are free because they should be.
# If they've saved you time or helped your team, and you want to
# support continued development, check out:
#
# DisplayCache: church digital signage that integrates with TouchPoint
# https://displaycache.com
#
# TPxi Go: your church contacts, wherever you work.
# Look up anyone in TouchPoint, log calls and emails from Outlook
# or your phone. No tab switching, no lost context.
# https://tpxigo.com
# ----------------------------------------------------------------
#
# TPxi Duplicates: find people who are in TouchPoint twice, check them, and
# merge them with TouchPoint's own person merge.
#
# WHO CAN USE IT (checked on the server for the page and every action, since
# /PyScriptForm does not enforce the #Roles line): Admin and Manager2, the
# roles TouchPoint's own merge screen lets delete a duplicate. Updating the
# script is Admin only. Giving (how many gifts, when) shows only to people
# with the Finance role: TouchPoint runs script SQL as a login that cannot
# read gifts otherwise. TouchPoint's merge moves every gift either way.
#
# UPDATES: an Admin sees a banner when a newer TPxi Duplicates is out, and
# Update saves it over this script under whatever name it was saved as.
# Saved marks (Not the same) are extra values, so they are untouched.
#####################################################################
# Finds people who are probably in TouchPoint twice, compares the two records,
# merges them with TouchPoint's own merge (under the admin's own sign-in, so it
# is in TouchPoint's merge history), and checks the result. It never merges by
# itself; an admin presses Merge for each pair.
#
# FEATURES
#   Finding pairs
#     - One pass over everyone living. Strong: same birthday and names alike
#       (nicknames and misspellings too). Name changed: same birthday and first
#       name, a different last name, and a shared email, cell or address. No
#       birthday: names alike plus a shared email or cell. Leaves out different
#       genders, Jr. and Sr., and pairs marked Not the same.
#     - Set aside in their own groups so they are not merged by mistake:
#       possibly twins, possibly parent and child (same name, one added as an
#       adult while the other was a child), and records a form wrote someone
#       else over (first name changed with gender or birthday in one save).
#     - Tags for a January 1 birthday and for two children who could be twins.
#     - Where duplicates come from (how and by whom each newer record was
#       added, as a rate) and who keeps getting duplicated.
#   Side by side
#     - What each record holds, with rows that open to show both: photo,
#       attendance (and meetings both were at), involvements, notes the viewer's
#       roles allow, tags, logins, extra values, medical info, and a timeline of
#       changes with a sentence on what happened.
#     - The original (older) record stays. When the newer one is the one in the
#       family, the original moves into that family, bringing its own children
#       if chosen. A child is not moved in with grandparents by default.
#     - Both families, with anyone else who looks like they are in both (a
#       birthday a digit off or missing too) and a Compare button.
#     - More than two records: flagged, and the result offers the next pair.
#   Choosing fields (read from TouchPoint's own merge screen)
#     - The newer cell and email, but never an address that bounced, was
#       blocked or reported spam in the last year. A phone typo is one digit off.
#     - The family's spelling, else the usual one when the other is rare.
#     - Last name: the family's when the family shares it, else the newer
#       record's (a name change), or the family's spelling when the newer one
#       misspells it. Married into the family being joined: last name, title
#       and marital status come from that record. The older record's dropped
#       name goes into Former Name, and a nickname into Goes By (Nick Name).
#     - Values TouchPoint or a kiosk fills in by itself (Just Added, Unknown,
#       not specified, "Added during online registration") give way to real
#       ones. Origin and entry point stay the original's; member status follows
#       the membership history; any text field can take a typed value.
#   Around TouchPoint's merge (what it would otherwise lose)
#     - Re-adds the duplicate's tags and rebuilds the tags it owned.
#     - Puts back family extra values, which TouchPoint deletes along with
#       anyone in the family.
#     - Settles extra values both records have, so no Field_mv copies are left.
#     - Carries over the check-in note and the MinistrySafe account link, which
#       are not on TouchPoint's merge screen.
#     - Warns when TouchPoint will replace medical and emergency info, and lists
#       what is gone for good when the record has it (past absences, pickup
#       permissions, directory privacy, e-sign, Process Builder, coupons).
#     - Afterward: checks that gifts, attendance, involvements, logins, checks
#       and notes all landed, reports a changed name, and moves people into the
#       family chosen (TouchPoint's own Person > Family > Add).
#   Merge log
#     - Each merge keeps, on the person kept: when, who, what the deleted record
#       held (name, birthday, contact, address, family, membership dates), the
#       choices made, and what happened or needs checking. Each Not the same
#       keeps who and when, on both people. Read it under Merge log on the list.
#     - Seen only by Admin and Manager2: the extra value is registered as a
#       standard extra value with those Visibility Roles (StandardExtraValues2),
#       which an Admin opening the page sets up; until then nothing is logged.
#     - Never includes giving: no gift counts, statements, envelopes or cards.
#
# INSTALL: paste as Special Content > Python named TPxi_Duplicates, then open
# /PyScriptForm/TPxi_Duplicates.
#
# VERSION HISTORY
#   1.0.0  2026-11-19  Initial release.
#####################################################################
import re
import datetime

VERSION = '1.0.0'
# Updates come from the same TPxi script service as the other TPxi tools: the page asks
# DC_API_BASE which version is newest, and an Admin's Update pulls the new code from
# DC_API_WORKER and saves it over this script.
DC_SCRIPT_ID = 'TPxi_Duplicates'
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'
SCRIPT = 'TPxi_Duplicates'
ROLES = ('Admin', 'Manager2')   # the roles TouchPoint's own merge screen offers Move and Delete to
def _in_role(r):
    try:
        return bool(model.UserIsInRole(r))
    except Exception:
        return False


# TouchPoint runs script SQL as a read-only login that is denied the Contribution table
# unless the person has the Finance role, so a gift count would make the query fail.
FINANCE = _in_role('Finance')


def gift_sql(sql):
    """Giving columns for people with the Finance role, empty (null) for everyone else."""
    if FINANCE:
        return (sql.replace('GIFTS_P', '(SELECT COUNT(*) FROM dbo.Contribution c WITH (NOLOCK) WHERE c.PeopleId = p.PeopleId)')
                   .replace('GIFTS_X', '(SELECT COUNT(*) FROM dbo.Contribution x WITH (NOLOCK) WHERE x.PeopleId = %(p)d)')
                   .replace('LASTGIFT_X', '(SELECT CONVERT(varchar(10), MAX(ContributionDate), 120) FROM dbo.Contribution x WITH (NOLOCK) WHERE x.PeopleId = %(p)d)')
                   .replace('GIFTS_KL', '(SELECT COUNT(*) FROM dbo.Contribution WITH (NOLOCK) WHERE PeopleId IN (%(k)d, %(l)d))'))
    return (sql.replace('GIFTS_P', 'CAST(NULL AS int)').replace('GIFTS_X', 'CAST(NULL AS int)')
               .replace('LASTGIFT_X', 'CAST(NULL AS varchar(10))').replace('GIFTS_KL', 'CAST(NULL AS int)'))


NOT_SAME = 'TPxiMergeNotSame'   # person extra value: ",123,456," the people this one is not (keeps its first name so saved marks still count)


# ---------------------------------------------------------------- JSON (IronPython-safe)
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
        return '{' + ','.join(_json_escape_string(unicode(k)) + ':' + _json_encode(v) for k, v in obj.items()) + '}'
    if isinstance(obj, (list, tuple)):
        return '[' + ','.join(_json_encode(x) for x in obj) + ']'
    return _json_escape_string(unicode(obj))


def safe_json(obj):
    try:
        return _json_encode(obj)
    except Exception as e:
        return '{"success":false,"message":' + _json_escape_string('safe_json: ' + repr(e)) + '}'


def load_json(text, default):
    import json
    try:
        return json.loads(text) if text else default
    except Exception:
        return default


# ---------------------------------------------------------------- helpers
def _param(name, default=''):
    try:
        v = getattr(model.Data, 'mg_' + name)
    except Exception:
        v = None
    if v is None or unicode(v) == '':
        return default
    return unicode(v)


def esc(s):
    return (unicode(s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;').replace("'", '&#39;'))


def to_int(v, default=0):
    try:
        return int(v)
    except Exception:
        return default


def can_use():
    for r in ROLES:
        try:
            if model.UserIsInRole(r):
                return True
        except Exception:
            pass
    return False


def s(v):
    """A database value as text: None is '', a .NET date is left to SQL (CONVERT)."""
    return unicode(v).strip() if v is not None else ''


def pair_ids():
    keep, lose = to_int(_param('keep')), to_int(_param('lose'))
    if not keep or not lose or keep == lose:
        raise Exception('Pick two different people.')
    return keep, lose


# ---------------------------------------------------------------- finding likely duplicates
# Everyone living is compared in one pass. Two people are a pair when:
#   strong   same full birthday, first names alike and last names alike
#   name     same full birthday, first names alike, a different last name, and a
#            shared email, cell or address, or a maiden or hyphenated name that links them
#   check    a birthday missing on one side, first and last names alike, and a shared
#            email or cell
# First names are alike when they are the same, one is the other's nickname (the
# nickname field, or a pair from this church's own data in TouchPoint's FirstNick),
# they sound alike and start with the same letter, or one starts the other (Chris,
# Christopher). Last names are alike when they are the same or sound alike.
# Left out: different genders, a different suffix (Jr. and Sr.), and pairs marked
# Not the same here.
FIND_SQL = """
SET NOCOUNT ON;
SELECT p.PeopleId, p.FamilyId,
  LOWER(LTRIM(RTRIM(ISNULL(p.FirstName, '')))) Fn, LOWER(LTRIM(RTRIM(ISNULL(p.NickName, '')))) Nk,
  LOWER(LTRIM(RTRIM(ISNULL(p.LastName, '')))) Ln, LOWER(LTRIM(RTRIM(ISNULL(p.MaidenName, '')))) Mn,
  ISNULL(p.GenderId, 0) G, LOWER(LTRIM(RTRIM(ISNULL(p.SuffixCode, '')))) Sfx,
  CASE WHEN p.BirthYear > 1900 AND p.BirthMonth BETWEEN 1 AND 12 AND p.BirthDay BETWEEN 1 AND 31
       THEN CAST(p.BirthYear AS varchar(4)) + '-' + RIGHT('0' + CAST(p.BirthMonth AS varchar(2)), 2) + '-' + RIGHT('0' + CAST(p.BirthDay AS varchar(2)), 2)
       ELSE '' END Dob,
  LOWER(LTRIM(RTRIM(ISNULL(p.EmailAddress, '')))) Em,
  RIGHT(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(p.CellPhone, ''), '-', ''), '(', ''), ')', ''), ' ', ''), '.', ''), 10) Cell,
  LOWER(LTRIM(RTRIM(ISNULL(f.AddressLineOne, '')))) + '|' + LEFT(ISNULL(f.ZipCode, ''), 5) Addr,
  p.CreatedDate Cr
INTO #p
FROM dbo.People p WITH (NOLOCK) JOIN dbo.Families f WITH (NOLOCK) ON f.FamilyId = p.FamilyId
WHERE ISNULL(p.IsDeceased, 0) = 0 AND ISNULL(p.ArchivedFlag, 0) = 0 AND ISNULL(p.FirstName, '') <> '';
CREATE INDEX ixd ON #p(Dob); CREATE INDEX ixe ON #p(Em); CREATE INDEX ixc ON #p(Cell);

SELECT LOWER(FirstName) A, LOWER(NickName) B INTO #nk FROM dbo.FirstNick WITH (NOLOCK) WHERE [count] >= 2;
CREATE INDEX ixn ON #nk(A, B);

SELECT a.PeopleId PA, b.PeopleId PB INTO #pr FROM #p a JOIN #p b ON a.Dob = b.Dob AND a.Dob <> '' AND a.PeopleId < b.PeopleId;
INSERT #pr SELECT a.PeopleId, b.PeopleId FROM #p a JOIN #p b ON a.Em = b.Em AND a.PeopleId < b.PeopleId
  WHERE a.Em LIKE '%@%' AND (a.Dob = '' OR b.Dob = '');
INSERT #pr SELECT a.PeopleId, b.PeopleId FROM #p a JOIN #p b ON a.Cell = b.Cell AND a.PeopleId < b.PeopleId
  WHERE LEN(a.Cell) = 10 AND (a.Dob = '' OR b.Dob = '');

SELECT DISTINCT x.PA, x.PB,
  CASE WHEN a.Dob <> '' AND a.Dob = b.Dob THEN 1 ELSE 0 END SBd,
  CASE WHEN a.Dob = '' OR b.Dob = '' THEN 1 ELSE 0 END NoBd,
  CASE WHEN a.Fn = b.Fn THEN 1 ELSE 0 END FSame,
  CASE WHEN a.Fn = b.Nk OR a.Nk = b.Fn OR (a.Nk <> '' AND a.Nk = b.Nk)
         OR EXISTS (SELECT 1 FROM #nk n WHERE (n.A = a.Fn AND n.B = b.Fn) OR (n.A = b.Fn AND n.B = a.Fn)) THEN 1 ELSE 0 END FNick,
  CASE WHEN SOUNDEX(a.Fn) = SOUNDEX(b.Fn) AND LEFT(a.Fn, 1) = LEFT(b.Fn, 1) THEN 1 ELSE 0 END FSound,
  CASE WHEN LEN(a.Fn) >= 3 AND LEN(b.Fn) >= 3 AND (a.Fn LIKE b.Fn + '%' OR b.Fn LIKE a.Fn + '%') THEN 1 ELSE 0 END FStart,
  CASE WHEN a.Ln = b.Ln THEN 1 ELSE 0 END LSame,
  CASE WHEN a.Ln <> b.Ln AND SOUNDEX(a.Ln) = SOUNDEX(b.Ln) THEN 1 ELSE 0 END LSound,
  CASE WHEN (a.Mn <> '' AND a.Mn = b.Ln) OR (b.Mn <> '' AND b.Mn = a.Ln) THEN 1 ELSE 0 END LMaiden,
  CASE WHEN a.Ln <> b.Ln AND LEN(a.Ln) >= 3 AND LEN(b.Ln) >= 3 AND (a.Ln LIKE '%' + b.Ln + '%' OR b.Ln LIKE '%' + a.Ln + '%') THEN 1 ELSE 0 END LHyphen,
  CASE WHEN a.Em LIKE '%@%' AND a.Em = b.Em THEN 1 ELSE 0 END SEm,
  CASE WHEN LEN(a.Cell) = 10 AND a.Cell = b.Cell THEN 1 ELSE 0 END SCell,
  CASE WHEN LEN(a.Addr) > 8 AND a.Addr = b.Addr THEN 1 ELSE 0 END SAddr,
  CASE WHEN a.FamilyId = b.FamilyId THEN 1 ELSE 0 END SFam,
  CASE WHEN a.G IN (1, 2) AND b.G IN (1, 2) AND a.G <> b.G THEN 1 ELSE 0 END GClash,
  CASE WHEN a.Sfx <> b.Sfx THEN 1 ELSE 0 END SfxClash,
  CASE WHEN a.Cr > b.Cr THEN a.Cr ELSE b.Cr END Newer,
  CASE WHEN a.Dob <> '' AND a.Dob = b.Dob AND RIGHT(a.Dob, 5) = '01-01' THEN 1 ELSE 0 END Jan1
INTO #s
FROM #pr x JOIN #p a ON a.PeopleId = x.PA JOIN #p b ON b.PeopleId = x.PB;

SELECT TOP 400 PA, PB, SBd, NoBd, FSame, FNick, FSound, FStart, LSame, LSound, LMaiden, LHyphen, SEm, SCell, SAddr, SFam, Jan1,
  CONVERT(varchar(10), Newer, 120) Newer,
  CASE WHEN SBd = 1 AND (LSame = 1 OR LSound = 1) THEN 'strong'
       WHEN SBd = 1 THEN 'name'
       ELSE 'check' END Tier
FROM #s
WHERE GClash = 0 AND SfxClash = 0
  AND (FSame = 1 OR FNick = 1 OR FSound = 1 OR FStart = 1)
  AND (
       (SBd = 1 AND (LSame = 1 OR LSound = 1))
    OR (SBd = 1 AND LSame = 0 AND LSound = 0 AND (SEm + SCell + SAddr > 0 OR LMaiden = 1 OR LHyphen = 1))
    OR (NoBd = 1 AND (LSame = 1 OR LSound = 1 OR LMaiden = 1 OR LHyphen = 1) AND (SEm = 1 OR SCell = 1))
  )
  AND {DAYS}
  AND NOT EXISTS (SELECT 1 FROM dbo.PeopleExtra e WITH (NOLOCK) WHERE e.PeopleId = PA AND e.Field = '{NOTSAME}' AND e.Data LIKE '%,' + CAST(PB AS varchar(12)) + ',%')
  AND NOT EXISTS (SELECT 1 FROM dbo.PeopleExtra e WITH (NOLOCK) WHERE e.PeopleId = PB AND e.Field = '{NOTSAME}' AND e.Data LIKE '%,' + CAST(PA AS varchar(12)) + ',%')
ORDER BY CASE WHEN SBd = 1 AND (LSame = 1 OR LSound = 1) THEN 0 WHEN SBd = 1 THEN 1 ELSE 2 END, Newer DESC
"""


def reasons_of(r):
    out = []
    if r.SBd:
        out.append('Same birthday')
    if r.NoBd:
        out.append('No birthday on one')
    if r.FSame and r.LSame:
        out.append('Same name')
    else:
        if r.FSame:
            out.append('Same first name')
        elif r.FNick:
            out.append('Nickname')
        elif r.FSound:
            out.append('First names sound alike')
        elif r.FStart:
            out.append('One first name starts the other')
        if r.LSame:
            out.append('Same last name')
        elif r.LSound:
            out.append('Last names sound alike')
        elif r.LMaiden:
            out.append('Maiden name matches')
        elif r.LHyphen:
            out.append('Hyphenated last name')
        else:
            out.append('Last name changed')
    if r.SEm:
        out.append('Same email')
    if r.SCell:
        out.append('Same cell')
    if r.SAddr:
        out.append('Same address')
    if r.SFam:
        out.append('Same family')
    if r.Jan1:
        out.append('Birthday is January 1 (often a stand-in)')
    return out


def clean_how(*vals):
    """How a record was added (origin, entry point), leaving out TouchPoint's 'not specified'."""
    out = []
    for v in vals:
        v = s(v)
        if v and not re.search(r'not\s*specified|^none$', v, re.I) and v not in out:
            out.append(v)
    return ', '.join(out)


def people_brief(ids):
    """Name, age, when and how added, and how much each record holds, to suggest who stays.
    The record in the family (head of household, with the spouse and children) counts most:
    TouchPoint's merge never moves anyone between families, so the person kept stays in
    the kept record's family."""
    out = {}
    if not ids:
        return out
    for r in q.QuerySql(gift_sql("""
            SELECT p.PeopleId, p.Name, p.Age, CONVERT(varchar(10), p.CreatedDate, 120) Created,
              DATEDIFF(year, p.CreatedDate, GETDATE()) Years,
              ISNULL(o.Description, '') Origin, ISNULL(ep.Description, '') EntryPoint, ISNULL(cb.Name, '') CreatedBy, ISNULL(p.CreatedBy, 0) CreatedById,
              (SELECT COUNT(*) FROM dbo.People x WITH (NOLOCK) WHERE x.FamilyId = p.FamilyId AND ISNULL(x.IsDeceased, 0) = 0) FamN,
              CASE WHEN f.HeadOfHouseholdId = p.PeopleId THEN 1 ELSE 0 END IsHead,
              GIFTS_P Gifts,
              (SELECT COUNT(*) FROM dbo.Attend a WITH (NOLOCK) WHERE a.PeopleId = p.PeopleId AND a.AttendanceFlag = 1) Att,
              (SELECT COUNT(*) FROM dbo.Users u WITH (NOLOCK) WHERE u.PeopleId = p.PeopleId) Logins,
              (SELECT COUNT(*) FROM dbo.OrganizationMembers om WITH (NOLOCK) WHERE om.PeopleId = p.PeopleId) Orgs,
              (SELECT COUNT(*) FROM dbo.TaskNote tn WITH (NOLOCK) WHERE tn.AboutPersonId = p.PeopleId) Notes,
              (SELECT COUNT(*) FROM dbo.TagPerson tp WITH (NOLOCK) JOIN dbo.Tag t WITH (NOLOCK) ON t.Id = tp.Id
                 WHERE tp.PeopleId = p.PeopleId AND t.TypeId = 1 AND t.StatusFlagId IS NULL) Tags,
              (SELECT COUNT(*) FROM dbo.MergeHistory mh WITH (NOLOCK) WHERE mh.ToId = p.PeopleId) Merged
            FROM dbo.People p WITH (NOLOCK)
            JOIN dbo.Families f WITH (NOLOCK) ON f.FamilyId = p.FamilyId
            LEFT JOIN lookup.Origin o ON o.Id = p.OriginId
            LEFT JOIN lookup.EntryPoint ep ON ep.Id = p.EntryPointId
            LEFT JOIN dbo.Users u ON u.UserId = p.CreatedBy
            LEFT JOIN dbo.People cb ON cb.PeopleId = u.PeopleId
            WHERE p.PeopleId IN (%s)""") % ','.join(str(i) for i in ids)):
        fam = max(0, to_int(r.FamN) - 1)
        hist = (min(to_int(r.Gifts), 50) * 3 + min(to_int(r.Att), 100) + min(to_int(r.Orgs), 10) * 5
                + min(to_int(r.Notes), 10) * 5 + min(to_int(r.Tags), 20) * 2 + to_int(r.Logins) * 20)
        # The original (older) record stays: its record number and added date never move, and its
        # membership dates only come along with its member status. Family placement is no longer a
        # reason to keep the newer one, since the person kept can be moved into the right family.
        score = [max(0, to_int(r.Years)), hist, 1 if fam else 0, fam]
        out[r.PeopleId] = {'pid': r.PeopleId, 'name': s(r.Name), 'age': r.Age, 'created': s(r.Created),
                           'how': clean_how(r.Origin, r.EntryPoint), 'createdBy': s(r.CreatedBy), 'createdById': to_int(r.CreatedById),
                           'gifts': to_int(r.Gifts) if FINANCE else None, 'att': to_int(r.Att), 'login': to_int(r.Logins) > 0,
                           'orgs': to_int(r.Orgs), 'family': fam, 'head': bool(r.IsHead),
                           'merged': to_int(r.Merged), 'score': score}
    return out


def act_find():
    days = to_int(_param('days'))
    cond = '1 = 1' if days <= 0 else "Newer >= DATEADD(day, -%d, GETDATE())" % max(1, min(3650, days))
    sql = FIND_SQL.replace('{DAYS}', cond).replace('{NOTSAME}', NOT_SAME)
    rows = list(q.QuerySql(sql))
    ids = set()
    for r in rows:
        ids.add(r.PA)
        ids.add(r.PB)
    brief = people_brief(list(ids))
    pairs = []
    for r in rows:
        a, b = brief.get(r.PA), brief.get(r.PB)
        if not a or not b:
            continue
        # the older record, then history, then the one in a family; then the lower id
        keep, lose = (a, b) if (a['score'], -a['pid']) >= (b['score'], -b['pid']) else (b, a)
        why = reasons_of(r)
        # two children with the same birthday and different first names (Brady and Brody): twins look like this
        if not r.FSame and not r.FNick and keep['age'] is not None and lose['age'] is not None and to_int(keep['age'], 99) < 18 and to_int(lose['age'], 99) < 18:
            why.append('Children: make sure they are not twins')
        pairs.append({'keep': keep, 'lose': lose, 'tier': s(r.Tier), 'why': why, 'newer': s(r.Newer),
                      'fsame': bool(r.FSame), 'sfam': bool(r.SFam)})
    # Possibly twins: first names that really differ (Alexa and Alyssa sound alike but are two
    # people) and either entered together (record numbers 1 to 3 apart, or one family added the
    # same day) or present at the same meeting more than once. A real duplicate usually has the
    # same first name (a typo in the last name, a family entered twice, or both records on a roll).
    if pairs:
        vals = ','.join('(%d,%d)' % (x['keep']['pid'], x['lose']['pid']) for x in pairs)
        extra = {}
        for r in q.QuerySql("""
                SELECT v.A, v.B, ABS(v.A - v.B) Gap,
                  CASE WHEN CONVERT(date, a.CreatedDate) = CONVERT(date, b.CreatedDate) THEN 1 ELSE 0 END SameDay,
                  CONVERT(varchar(10), a.CreatedDate, 120) DayA,
                  (SELECT COUNT(*) FROM dbo.Attend x WITH (NOLOCK) JOIN dbo.Attend y WITH (NOLOCK) ON y.MeetingId = x.MeetingId AND y.PeopleId = v.B AND y.AttendanceFlag = 1
                     WHERE x.PeopleId = v.A AND x.AttendanceFlag = 1) Co,
                  ISNULL(a.PositionInFamilyId, 0) PosA, ISNULL(b.PositionInFamilyId, 0) PosB,
                  CASE WHEN a.BirthYear > 1900 AND a.BirthMonth > 0 AND a.BirthDay > 0 THEN 1 ELSE 0 END DobA,
                  CASE WHEN b.BirthYear > 1900 AND b.BirthMonth > 0 AND b.BirthDay > 0 THEN 1 ELSE 0 END DobB,
                  CASE WHEN b.BirthYear > 1900 THEN YEAR(a.CreatedDate) - b.BirthYear END AgeBAtA,
                  CASE WHEN a.BirthYear > 1900 THEN YEAR(b.CreatedDate) - a.BirthYear END AgeAAtB,
                  YEAR(a.CreatedDate) YrA, YEAR(b.CreatedDate) YrB,
                  (SELECT COUNT(*) FROM dbo.OrganizationMembers x WITH (NOLOCK) JOIN dbo.OrganizationMembers y WITH (NOLOCK) ON y.OrganizationId = x.OrganizationId AND y.PeopleId = v.B
                     WHERE x.PeopleId = v.A) CommonOrgs
                FROM (VALUES %s) v(A, B) JOIN dbo.People a WITH (NOLOCK) ON a.PeopleId = v.A JOIN dbo.People b WITH (NOLOCK) ON b.PeopleId = v.B""" % vals):
            extra[(r.A, r.B)] = r
        for x in pairs:
            e = extra.get((x['keep']['pid'], x['lose']['pid']))
            together = bool(e) and (to_int(e.Gap) <= 3 or (x['sfam'] and to_int(e.SameDay)))
            co = to_int(e.Co) if e else 0
            # twins never share a first name: the same name present at the same meeting is one person
            # on the roll twice (a leader marked both records), not two people
            x['twins'] = (together or co >= 2) and not x['fsame']
            if x['twins']:
                x['twinWhy'] = ('added together' + (' on ' + s(e.DayA) if e and to_int(e.SameDay) else '') if together else '') + \
                    ((' and ' if together else '') + 'present together at %d meetings' % co if co >= 2 else '')
            # Possibly parent and child: one has no birthday and was added as an adult (head or spouse
            # of a family) when the other was a child, and they share no involvement. A parent and a
            # child with the same name share an email or phone (Amy and Amy Williamson).
            x['parent'] = False
            if e and not x['twins']:
                for adult, child, pos, dob_a, dob_c, age_c, yr in (('keep', 'lose', e.PosA, e.DobA, e.DobB, e.AgeBAtA, e.YrA),
                                                                   ('lose', 'keep', e.PosB, e.DobB, e.DobA, e.AgeAAtB, e.YrB)):
                    if to_int(pos) in (10, 20) and not to_int(dob_a) and to_int(dob_c) and age_c is not None and 0 <= to_int(age_c, 99) < 16 and not to_int(e.CommonOrgs):
                        x['parent'] = True
                        x['parentWhy'] = '#%d was added as an adult in %d, when #%d was %d' % (x[adult]['pid'], to_int(yr), x[child]['pid'], to_int(age_c))
                        break
    # Record overwritten: one of the two had its first name changed together with its gender or
    # birthday. That is usually someone else's record written over (a parent registering a child
    # while signed in), so the fix is to change it back, not to merge.
    ids = set()
    for x in pairs:
        ids.add(x['keep']['pid'])
        ids.add(x['lose']['pid'])
    ident = identity_changes(list(ids)) if ids else {}
    for x in pairs:
        for side in ('keep', 'lose'):
            ch = ident.get(x[side]['pid'])
            if ch:
                c = ch[-1]
                x['over'] = {'pid': x[side]['pid'], 'from': c['from'], 'to': c['to'], 'when': c['when']}
    counts = {'strong': 0, 'name': 0, 'check': 0, 'twins': 0, 'over': 0, 'parent': 0}
    for p in pairs:
        if p.get('over'):
            counts['over'] += 1
        elif p.get('twins'):
            counts['twins'] += 1
        elif p.get('parent'):
            counts['parent'] += 1
        else:
            counts[p['tier']] = counts.get(p['tier'], 0) + 1
    # where duplicates come from: how the newer record of each pair was added
    how, by = {}, {}
    for p in pairs:
        if p.get('twins') or p.get('over') or p.get('parent'):
            continue
        n = p['lose'] if p['lose']['created'] >= p['keep']['created'] else p['keep']
        h = n['how'] or 'Not recorded'
        how[h] = how.get(h, 0) + 1
        if n['createdBy'] and n['createdById']:
            key = n['createdById']
            if key not in by:
                by[key] = [n['createdBy'], 0, 0]
            by[key][1] += 1
    # the fair number is the rate: duplicates out of everyone that person added
    if by:
        for r in q.QuerySql('SELECT CreatedBy U, COUNT(*) N FROM dbo.People WITH (NOLOCK) WHERE CreatedBy IN (%s) GROUP BY CreatedBy' % ','.join(str(i) for i in by)):
            if to_int(r.U) in by:
                by[to_int(r.U)][2] = to_int(r.N)
    # people merged more than once: something keeps adding them again
    repeats = [{'pid': r.ToId, 'name': s(r.Name), 'n': to_int(r.N), 'last': s(r.LastDt), 'first': s(r.FirstDt)} for r in q.QuerySql("""
        SELECT TOP 40 m.ToId, p.Name, COUNT(*) N, CONVERT(varchar(10), MAX(m.Dt), 120) LastDt, CONVERT(varchar(10), MIN(m.Dt), 120) FirstDt
        FROM dbo.MergeHistory m WITH (NOLOCK) JOIN dbo.People p WITH (NOLOCK) ON p.PeopleId = m.ToId
        WHERE m.Action LIKE '%Delete%'
        GROUP BY m.ToId, p.Name HAVING COUNT(*) >= 2 ORDER BY COUNT(*) DESC, MAX(m.Dt) DESC""")]
    return {'success': True, 'pairs': pairs, 'counts': counts, 'capped': len(rows) >= 400,
            'how': sorted([[k, v] for k, v in how.items()], key=lambda x: -x[1]),
            'by': sorted([[v[0], v[1], v[2]] for v in by.values()], key=lambda x: -x[1])[:8], 'repeats': repeats}


def act_lookup():
    """Find people by name or id, for comparing any two by hand."""
    t = _param('q').strip()
    if not t:
        return {'success': True, 'people': []}
    if re.match(r'^\d+$', t):
        where = 'p.PeopleId = %d' % to_int(t)
    else:
        parts = [x.replace("'", "''") for x in t.split() if x][:3]
        conds = ["(p.FirstName LIKE '{0}%' OR p.NickName LIKE '{0}%' OR p.LastName LIKE '{0}%')".format(x) for x in parts]
        where = ' AND '.join(conds) if conds else '1 = 0'
    rows = q.QuerySql("""SELECT TOP 12 p.PeopleId, p.Name, p.Age, CONVERT(varchar(10), p.CreatedDate, 120) Created
        FROM dbo.People p WITH (NOLOCK) WHERE %s AND ISNULL(p.IsDeceased, 0) = 0 ORDER BY p.Name2""" % where)
    return {'success': True, 'people': [{'pid': r.PeopleId, 'name': s(r.Name), 'age': r.Age, 'created': s(r.Created)} for r in rows]}


# ---------------------------------------------------------------- one pair, side by side
IDENT_FIELDS = ('FirstName', 'LastName', 'NickName', 'GenderId', 'BirthDay', 'BirthMonth', 'BirthYear', 'MaritalStatusId')
GENDER = {'1': 'Male', '2': 'Female', '0': 'Unknown'}


def login_before(pid, at):
    """The sign-in to this record's login in the hour before a change, as 'HH:MM as username'."""
    try:
        r = q.QuerySqlTop1("""SELECT TOP 1 CONVERT(varchar(5), a.ActivityDate, 108) T, ISNULL(u.Username, '') N FROM dbo.ActivityLog a WITH (NOLOCK)
            JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = a.UserId
            WHERE u.PeopleId = %d AND a.Activity LIKE '%%logged in%%' AND a.ActivityDate BETWEEN DATEADD(minute, -60, '%s') AND '%s'
            ORDER BY a.ActivityDate DESC""" % (pid, at, at))
        return (s(r.T) + ' as ' + s(r.N)) if r else ''
    except Exception:
        return ''


def identity_changes(pids):
    """Saves that changed a record's first name together with its gender or birthday: usually a
    different person written over this record (a parent signed in and registering a child, whose
    answers the form saved onto the parent). Returns {pid: [{when, by, self, changes: [...]}]}."""
    out = {}
    if not pids:
        return out
    rows = q.QuerySql("""
        SELECT c.Id, c.PeopleId, ISNULL(c.UserPeopleId, 0) U, CONVERT(varchar(10), c.Created, 120) D, CONVERT(varchar(19), c.Created, 120) T, ISNULL(u.Name, '') ByName,
          d.Field, ISNULL(d.Before, '') B, ISNULL(d.After, '') A
        FROM dbo.ChangeLog c WITH (NOLOCK)
        JOIN dbo.ChangeDetails d WITH (NOLOCK) ON d.Id = c.Id
        LEFT JOIN dbo.People u WITH (NOLOCK) ON u.PeopleId = c.UserPeopleId
        WHERE c.PeopleId IN (%s) AND d.Field IN ('FirstName', 'LastName', 'NickName', 'GenderId', 'BirthDay', 'BirthMonth', 'BirthYear', 'MaritalStatusId')
          AND c.Id IN (SELECT f.Id FROM dbo.ChangeDetails f WITH (NOLOCK) WHERE f.Field = 'FirstName' AND ISNULL(f.Before, '') NOT IN ('', '(null)'))
        ORDER BY c.Created""" % ','.join(str(i) for i in pids))
    saves = {}
    for r in rows:
        key = (r.PeopleId, r.Id)
        if key not in saves:
            saves[key] = {'pid': r.PeopleId, 'when': s(r.D), 'at': s(r.T), 'by': s(r.ByName), 'self': to_int(r.U) == r.PeopleId, 'f': {}}
        saves[key]['f'][s(r.Field)] = (s(r.B), s(r.A))
    for key in sorted(saves, key=lambda k: saves[k]['when']):
        sv = saves[key]
        f = sv['f']
        if not ('GenderId' in f and f['GenderId'][0] not in ('', '(null)', '0')) and not any(x in f for x in ('BirthDay', 'BirthMonth', 'BirthYear')):
            continue   # a first name fixed by itself is a correction, not a different person
        def bd(i):
            parts = [f.get(x, ('', ''))[i] for x in ('BirthMonth', 'BirthDay', 'BirthYear')]
            parts = [x for x in parts if x and x != '(null)']
            return '/'.join(parts)
        # the registration that saved it, if one finished within 15 minutes for this record
        reg = q.QuerySqlTop1("""SELECT TOP 1 o.OrganizationName O FROM dbo.RegPeople rp WITH (NOLOCK)
            JOIN dbo.Registration rg WITH (NOLOCK) ON rg.RegistrationId = rp.RegistrationId
            JOIN dbo.Organizations o WITH (NOLOCK) ON o.OrganizationId = rg.OrganizationId
            WHERE rp.PeopleId = %d AND rg.CompletedDate BETWEEN DATEADD(minute, -15, '%s') AND DATEADD(minute, 15, '%s')""" % (sv['pid'], sv['at'], sv['at']))
        g0 = GENDER.get(f.get('GenderId', ('', ''))[0], '')
        g1 = GENDER.get(f.get('GenderId', ('', ''))[1], '')
        out.setdefault(sv['pid'], []).append({
            'when': sv['when'], 'by': sv['by'], 'self': sv['self'],
            'from': f['FirstName'][0], 'to': f['FirstName'][1],
            'genderFrom': g0, 'genderTo': g1, 'birthTo': bd(1), 'birthFrom': bd(0), 'reg': s(reg.O) if reg else '',
            'regAt': '', 'at': sv['at'][11:16], 'login': login_before(sv['pid'], sv['at'])})
    return out


FIELD_LABELS = {'FirstName': 'first name', 'LastName': 'last name', 'NickName': 'goes by', 'MiddleName': 'middle name',
                'GenderId': 'gender', 'MaritalStatusId': 'marital status', 'EmailAddress': 'email', 'CellPhone': 'cell',
                'MemberStatusId': 'member status', 'ArchivedFlag': 'archived', 'PositionInFamilyId': 'position in family',
                'TitleCode': 'title', 'HomePhone': 'home phone', 'AddressLineOne': 'address', 'DeceasedDate': 'deceased date',
                'GradeLevel': 'grade', 'CampusId': 'campus', 'EnvelopeOptionsId': 'envelope option', 'ContributionOptionsId': 'statement option',
                'EntryPointId': 'entry point', 'OriginId': 'origin', 'WorkPhone': 'work phone', 'EmailAddress2': 'email 2',
                'BaptismStatusId': 'baptism status', 'BaptismTypeId': 'baptism type', 'DecisionTypeId': 'decision', 'NewMemberClassStatusId': 'new member class',
                'JoinCodeId': 'joined by', 'GradeLevelId': 'grade level', 'ElectronicStatement': 'electronic statement', 'Grade': 'school grade', 'DropCodeId': 'dropped by', 'LetterStatusId': 'letter status', 'DOB': 'birthday', 'MaidenName': 'maiden name'}


def _lookup(table):
    try:
        return dict((str(r.Id), s(r.Description)) for r in q.QuerySql('SELECT Id, Description FROM lookup.%s' % table))
    except Exception:
        return {}


def _people_names(ids):
    ids = [i for i in set(ids) if i]
    if not ids:
        return {}
    return dict((r.PeopleId, s(r.Name)) for r in q.QuerySql('SELECT PeopleId, Name FROM dbo.People WITH (NOLOCK) WHERE PeopleId IN (%s)' % ','.join(str(i) for i in ids)))


def timeline(pid):
    """What happened to one record, oldest first: added, saves (with before and after), registrations
    and the name each was entered under, sign-ins on the days of a registration or a name change,
    merges, archive, and first and last attendance and gift."""
    ev = []
    p = q.QuerySqlTop1("""SELECT CONVERT(varchar(16), p.CreatedDate, 120) C, ISNULL(cb.Name, '') ByName, ISNULL(o.Description, '') O, ISNULL(ep.Description, '') E
        FROM dbo.People p WITH (NOLOCK) LEFT JOIN dbo.Users u ON u.UserId = p.CreatedBy LEFT JOIN dbo.People cb ON cb.PeopleId = u.PeopleId
        LEFT JOIN lookup.Origin o ON o.Id = p.OriginId LEFT JOIN lookup.EntryPoint ep ON ep.Id = p.EntryPointId WHERE p.PeopleId = %d""" % pid)
    if p:
        how = clean_how(p.O, p.E)
        ev.append({'at': s(p.C), 'kind': 'added', 'text': 'Record added' + (' by ' + s(p.ByName) if s(p.ByName) else '') + (' (' + how + ')' if how else '')})
    look = {'MaritalStatusId': _lookup('MaritalStatus'), 'MemberStatusId': _lookup('MemberStatus'), 'PositionInFamilyId': _lookup('FamilyPosition'),
            'GenderId': {'0': 'Unknown', '1': 'Male', '2': 'Female'}, 'ArchivedFlag': {'True': 'yes', 'False': 'no'}}
    for fld, table in (('JoinCodeId', 'JoinType'), ('DropCodeId', 'DropType'), ('BaptismStatusId', 'BaptismStatus'), ('BaptismTypeId', 'BaptismType'),
                       ('DecisionTypeId', 'DecisionType'), ('NewMemberClassStatusId', 'NewMemberClassStatus'), ('EnvelopeOptionsId', 'EnvelopeOption'),
                       ('ContributionOptionsId', 'EnvelopeOption'), ('CampusId', 'Campus'), ('LetterStatusId', 'MemberLetterStatus'),
                       ('OriginId', 'Origin'), ('EntryPointId', 'EntryPoint'), ('GradeLevelId', 'GradeLevel')):
        look[fld] = _lookup(table)
    groups, order, users = {}, [], []
    for r in q.QuerySql("""SELECT TOP 500 CONVERT(varchar(16), Created, 120) T, ISNULL(UserPeopleId, 0) U, Field, ISNULL(Before, '') B, ISNULL(After, '') A
            FROM dbo.ChangeLogDetails WITH (NOLOCK) WHERE PeopleId = %d ORDER BY Created""" % pid):
        key = (s(r.T), to_int(r.U))
        if key not in groups:
            groups[key] = {}
            order.append(key)
            users.append(to_int(r.U))
        groups[key][s(r.Field)] = (s(r.B), s(r.A))
    names = _people_names(users)
    days = set()

    def val(f, v):
        v = '' if v in ('(null)', 'None') else v
        return look.get(f, {}).get(v, v) or 'none'
    for key in order:
        f = groups[key]
        parts = []
        if any(x in f for x in ('BirthMonth', 'BirthDay', 'BirthYear')):
            new = '/'.join(x for x in (f.get('BirthMonth', ('', ''))[1], f.get('BirthDay', ('', ''))[1], f.get('BirthYear', ('', ''))[1]) if x and x != '(null)')
            parts.append('birthday set to ' + (new or 'none'))
        for fld in sorted(f):
            if fld.startswith('Birth'):
                continue
            b, a = f[fld]
            if fld == 'ArchivedFlag':
                parts.append('archived' if a == 'True' else 'unarchived')
                continue
            if b in ('', '(null)', 'None') and a in ('0', ''):
                continue   # an empty code set to 0: TouchPoint filling in, not a change worth reading
            parts.append('%s %s to %s' % (FIELD_LABELS.get(fld, fld), val(fld, b), val(fld, a)))
        ident = 'FirstName' in f and f['FirstName'][0] not in ('', '(null)') and ('GenderId' in f or any(x.startswith('Birth') for x in f))
        who = 'by themselves (signed in)' if key[1] == pid else ('by ' + names.get(key[1], '') if names.get(key[1]) else '')
        if not parts:
            continue
        ev.append({'at': key[0], 'kind': 'ident' if ident else 'change', 'text': 'Changed: ' + '; '.join(parts) + (' ' + who if who else '')})
        if ident:
            days.add(key[0][:10])
    for r in q.QuerySql("""SELECT CONVERT(varchar(16), rg.CompletedDate, 120) T, ISNULL(o.OrganizationName, '') O, ISNULL(rp.FirstName, '') F, ISNULL(rp.LastName, '') L
            FROM dbo.RegPeople rp WITH (NOLOCK) JOIN dbo.Registration rg WITH (NOLOCK) ON rg.RegistrationId = rp.RegistrationId
            LEFT JOIN dbo.Organizations o WITH (NOLOCK) ON o.OrganizationId = rg.OrganizationId
            WHERE rp.PeopleId = %d AND rg.CompletedDate IS NOT NULL ORDER BY rg.CompletedDate""" % pid):
        ev.append({'at': s(r.T), 'kind': 'reg', 'text': 'Registration finished for ' + s(r.O) + ', with this record entered as ' + (s(r.F) + ' ' + s(r.L)).strip()})
        days.add(s(r.T)[:10])
    # sign-ins, only on the days something above happened (a person may sign in hundreds of times)
    if days:
        for r in q.QuerySql("""SELECT TOP 40 CONVERT(varchar(16), a.ActivityDate, 120) T, ISNULL(u.Username, '') N FROM dbo.ActivityLog a WITH (NOLOCK)
                JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = a.UserId
                WHERE u.PeopleId = %d AND a.Activity LIKE '%%logged in%%' AND CONVERT(date, a.ActivityDate) IN (%s) ORDER BY a.ActivityDate""" % (pid, ','.join("'%s'" % d for d in sorted(days)))):
            ev.append({'at': s(r.T), 'kind': 'login', 'text': 'Signed in as ' + s(r.N)})
    for r in q.QuerySql("""SELECT CONVERT(varchar(16), Dt, 120) T, FromId, ISNULL(FromName, '') N, ISNULL(WhoName, '') W FROM dbo.MergeHistory WITH (NOLOCK)
            WHERE ToId = %d AND Action LIKE '%%Delete%%' ORDER BY Dt""" % pid):
        ev.append({'at': s(r.T), 'kind': 'merge', 'text': 'Merged in #%d (%s)%s' % (r.FromId, s(r.N), ' by ' + s(r.W) if s(r.W) else '')})
    a = q.QuerySqlTop1("""SELECT CONVERT(varchar(10), MIN(MeetingDate), 120) F, CONVERT(varchar(10), MAX(MeetingDate), 120) L, COUNT(*) N
        FROM dbo.Attend WITH (NOLOCK) WHERE PeopleId = %d AND AttendanceFlag = 1""" % pid)
    if a and to_int(a.N):
        ev.append({'at': s(a.F), 'kind': 'att', 'text': 'First attended'})
        if to_int(a.N) > 1:
            ev.append({'at': s(a.L), 'kind': 'att', 'text': 'Last attended (%d in all)' % to_int(a.N)})
    gft = q.QuerySqlTop1("""SELECT CONVERT(varchar(10), MIN(ContributionDate), 120) F, CONVERT(varchar(10), MAX(ContributionDate), 120) L, COUNT(*) N
        FROM dbo.Contribution WITH (NOLOCK) WHERE PeopleId = %d""" % pid) if FINANCE else None
    if gft and to_int(gft.N):
        ev.append({'at': s(gft.F), 'kind': 'gift', 'text': 'First gift'})
        if to_int(gft.N) > 1:
            ev.append({'at': s(gft.L), 'kind': 'gift', 'text': 'Last gift (%d in all)' % to_int(gft.N)})
    ev.sort(key=lambda e: e['at'])
    return ev[-80:]


def person_detail(pid):
    r = q.QuerySqlTop1("""
        SELECT p.PeopleId, p.Name, p.FirstName, ISNULL(p.NickName, '') Nick, ISNULL(p.LastName, '') LastName, p.Age,
          p.BirthMonth, p.BirthDay, p.BirthYear, ISNULL(p.EmailAddress, '') Email, ISNULL(p.CellPhone, '') Cell,
          ISNULL(g.Description, '') Gender, ISNULL(ms.Description, '') MemberStatus, ISNULL(c.Description, '') Campus,
          ISNULL(fp.Description, '') Position, p.FamilyId, ISNULL(p.PictureId, 0) Pic,
          CONVERT(varchar(10), p.CreatedDate, 120) Created, ISNULL(cb.Name, '') CreatedBy,
          ISNULL(o.Description, '') Origin, ISNULL(ep.Description, '') EntryPoint,
          CASE WHEN f.HeadOfHouseholdId = p.PeopleId THEN 1 ELSE 0 END IsHead,
          (SELECT COUNT(*) FROM dbo.People x WITH (NOLOCK) WHERE x.FamilyId = p.FamilyId) FamN
        FROM dbo.People p WITH (NOLOCK)
        JOIN dbo.Families f WITH (NOLOCK) ON f.FamilyId = p.FamilyId
        LEFT JOIN lookup.Gender g ON g.Id = p.GenderId
        LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
        LEFT JOIN lookup.Campus c ON c.Id = p.CampusId
        LEFT JOIN lookup.FamilyPosition fp ON fp.Id = p.PositionInFamilyId
        LEFT JOIN lookup.Origin o ON o.Id = p.OriginId
        LEFT JOIN lookup.EntryPoint ep ON ep.Id = p.EntryPointId
        LEFT JOIN dbo.Users u ON u.UserId = p.CreatedBy
        LEFT JOIN dbo.People cb ON cb.PeopleId = u.PeopleId
        WHERE p.PeopleId = %d""" % pid)
    if not r:
        return None
    bd = ''
    if r.BirthMonth and r.BirthDay:
        bd = '%d/%d' % (r.BirthMonth, r.BirthDay) + ('/%d' % r.BirthYear if r.BirthYear else '')
    d = {'pid': pid, 'name': s(r.Name), 'first': s(r.FirstName), 'nick': s(r.Nick), 'last': s(r.LastName), 'age': r.Age,
         'birthday': bd, 'email': s(r.Email), 'cell': s(r.Cell), 'gender': s(r.Gender), 'member': s(r.MemberStatus),
         'campus': s(r.Campus), 'position': s(r.Position), 'fid': r.FamilyId, 'photo': to_int(r.Pic) > 0,
         'created': s(r.Created), 'createdBy': s(r.CreatedBy), 'how': clean_how(r.Origin, r.EntryPoint),
         'isHead': bool(r.IsHead), 'familySize': to_int(r.FamN)}
    c = q.QuerySqlTop1(gift_sql("""
        SELECT
          GIFTS_X Gifts,
          LASTGIFT_X LastGift,
          (SELECT COUNT(*) FROM dbo.Attend x WITH (NOLOCK) WHERE x.PeopleId = %(p)d AND x.AttendanceFlag = 1) Att,
          (SELECT COUNT(*) FROM dbo.Attend x WITH (NOLOCK) WHERE x.PeopleId = %(p)d AND x.AttendanceFlag = 0) Absent,
          (SELECT CONVERT(varchar(10), MAX(MeetingDate), 120) FROM dbo.Attend x WITH (NOLOCK) WHERE x.PeopleId = %(p)d AND x.AttendanceFlag = 1) LastAtt,
          (SELECT COUNT(*) FROM dbo.OrganizationMembers x WITH (NOLOCK) WHERE x.PeopleId = %(p)d) Orgs,
          (SELECT COUNT(*) FROM dbo.Users x WITH (NOLOCK) WHERE x.PeopleId = %(p)d) Logins,
          (SELECT CONVERT(varchar(10), MAX(CASE WHEN ISNULL(LastLoginDate, '1900-01-01') > ISNULL(LastActivityDate, '1900-01-01') THEN LastLoginDate ELSE LastActivityDate END), 120)
             FROM dbo.Users x WITH (NOLOCK) WHERE x.PeopleId = %(p)d) LastLogin,
          (SELECT COUNT(*) FROM dbo.RegPeople x WITH (NOLOCK) WHERE x.PeopleId = %(p)d) Regs,
          (SELECT COUNT(*) FROM dbo.BackgroundChecks x WITH (NOLOCK) WHERE x.PeopleID = %(p)d) Checks,
          (SELECT COUNT(*) FROM dbo.TaskNote x WITH (NOLOCK) WHERE x.AboutPersonId = %(p)d) Notes,
          (SELECT COUNT(*) FROM dbo.TagPerson tp WITH (NOLOCK) JOIN dbo.Tag t WITH (NOLOCK) ON t.Id = tp.Id WHERE tp.PeopleId = %(p)d AND t.TypeId = 1 AND t.StatusFlagId IS NULL) Tags,
          (SELECT COUNT(*) FROM dbo.TagPerson tp WITH (NOLOCK) JOIN dbo.Tag t WITH (NOLOCK) ON t.Id = tp.Id WHERE tp.PeopleId = %(p)d AND t.StatusFlagId IS NOT NULL) Flags,
          (SELECT COUNT(*) FROM dbo.CardIdentifiers x WITH (NOLOCK) WHERE x.PeopleId = %(p)d) Cards""") % {'p': pid})
    d['counts'] = {'gifts': to_int(c.Gifts) if FINANCE else None, 'lastGift': s(c.LastGift), 'att': to_int(c.Att), 'absent': to_int(c.Absent),
                   'lastAtt': s(c.LastAtt), 'orgs': to_int(c.Orgs), 'logins': to_int(c.Logins), 'lastLogin': s(c.LastLogin),
                   'regs': to_int(c.Regs), 'checks': to_int(c.Checks), 'notes': to_int(c.Notes), 'tags': to_int(c.Tags),
                   'flags': to_int(c.Flags), 'cards': to_int(c.Cards)}
    d['orgs'] = [{'oid': x.OrganizationId, 'name': s(x.OrganizationName), 'as': s(x.MT)} for x in q.QuerySql("""
        SELECT TOP 8 o.OrganizationId, o.OrganizationName, ISNULL(mt.Description, '') MT FROM dbo.OrganizationMembers om WITH (NOLOCK)
        JOIN dbo.Organizations o WITH (NOLOCK) ON o.OrganizationId = om.OrganizationId
        LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
        WHERE om.PeopleId = %d ORDER BY om.EnrollmentDate DESC""" % pid)]
    m = q.QuerySqlTop1("""
        SELECT CONVERT(varchar(10), p.JoinDate, 120) JoinDate, ISNULL(jt.Description, '') JoinType,
          CONVERT(varchar(10), p.DecisionDate, 120) DecisionDate, ISNULL(dt.Description, '') DecisionType,
          CONVERT(varchar(10), p.BaptismDate, 120) BaptismDate, ISNULL(bt.Description, '') BaptismType,
          CONVERT(varchar(10), p.NewMemberClassDate, 120) NmcDate,
          CONVERT(varchar(10), p.DropDate, 120) DropDate, ISNULL(dr.Description, '') DropType
        FROM dbo.People p WITH (NOLOCK)
        LEFT JOIN lookup.JoinType jt ON jt.Id = p.JoinCodeId
        LEFT JOIN lookup.DecisionType dt ON dt.Id = p.DecisionTypeId
        LEFT JOIN lookup.BaptismType bt ON bt.Id = p.BaptismTypeId
        LEFT JOIN lookup.DropType dr ON dr.Id = p.DropCodeId
        WHERE p.PeopleId = %d""" % pid)
    mem = []
    if m:
        if s(m.JoinDate):
            mem.append('Joined ' + s(m.JoinDate) + (' (' + s(m.JoinType) + ')' if s(m.JoinType) else ''))
        if s(m.DecisionDate) or (s(m.DecisionType) and not re.search(r'unknown', s(m.DecisionType), re.I)):
            mem.append('Decision ' + (s(m.DecisionDate) + ' ' if s(m.DecisionDate) else '') + ('(' + s(m.DecisionType) + ')' if s(m.DecisionType) else ''))
        if s(m.BaptismDate):
            mem.append('Baptized ' + s(m.BaptismDate) + (' (' + s(m.BaptismType) + ')' if s(m.BaptismType) else ''))
        if s(m.NmcDate):
            mem.append('New member class ' + s(m.NmcDate))
        if s(m.DropDate):
            mem.append('Dropped ' + s(m.DropDate) + (' (' + s(m.DropType) + ')' if s(m.DropType) else ''))
    d['membership'] = mem
    d['identity'] = identity_changes([pid]).get(pid, [])
    d['merged'] = [{'from': x.FromId, 'name': s(x.FromName), 'when': s(x.D), 'who': s(x.WhoName)} for x in q.QuerySql("""
        SELECT TOP 10 FromId, FromName, CONVERT(varchar(10), Dt, 120) D, WhoName FROM dbo.MergeHistory WITH (NOLOCK)
        WHERE ToId = %d AND Action LIKE '%%Delete%%' ORDER BY Dt DESC""" % pid)]
    return d


def family_of(fid, skip):
    out = []
    for r in q.QuerySql("""
            SELECT p.PeopleId, p.Name, ISNULL(p.FirstName, '') FirstName, ISNULL(p.NickName, '') Nick, ISNULL(p.LastName, '') LastName,
              p.Age, p.BirthMonth, p.BirthDay, p.BirthYear, ISNULL(fp.Description, '') Position
            FROM dbo.People p WITH (NOLOCK) LEFT JOIN lookup.FamilyPosition fp ON fp.Id = p.PositionInFamilyId
            WHERE p.FamilyId = %d AND ISNULL(p.IsDeceased, 0) = 0 ORDER BY p.PositionInFamilyId, p.Age DESC""" % fid):
        out.append({'pid': r.PeopleId, 'name': s(r.Name), 'first': s(r.FirstName).lower(), 'nick': s(r.Nick).lower(), 'last': s(r.LastName).strip(),
                    'age': r.Age, 'bm': to_int(r.BirthMonth), 'bd': to_int(r.BirthDay), 'by': to_int(r.BirthYear), 'position': s(r.Position),
                    'self': r.PeopleId in skip})
    return out


def related_of(fid):
    return [{'rel': s(r.Rel), 'name': s(r.Head)} for r in q.QuerySql("""
        SELECT ISNULL(r.FamilyRelationshipDesc, '') Rel, ISNULL(h.Name, '') Head FROM dbo.RelatedFamilies r WITH (NOLOCK)
        JOIN dbo.Families f WITH (NOLOCK) ON f.FamilyId = CASE WHEN r.FamilyId = %(f)d THEN r.RelatedFamilyId ELSE r.FamilyId END
        LEFT JOIN dbo.People h WITH (NOLOCK) ON h.PeopleId = f.HeadOfHouseholdId
        WHERE r.FamilyId = %(f)d OR r.RelatedFamilyId = %(f)d""" % {'f': fid})]


def _lev(a, b):
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (0 if ca == cb else 1)))
        prev = cur
    return prev[-1]


def match_families(a, b):
    """Pair up the people in the two families (other than the two being compared). 'same': names
    alike (the same, a nickname, one starting the other, a letter off) and the same birthday: one
    family entered twice (Lucy and Lucille). 'maybe': the same first name, and a birthday off by a
    digit in the same year or missing on one (Baker Blackman 9/8/2016 and 8/6/2016; Esseri Harris
    and Esseri Holmes), which the main list cannot find."""
    pairs = []
    used = set()
    for x in a:
        if x['self']:
            continue
        xn = set(n for n in (x['first'], x['nick']) if n)
        for y in b:
            if y['self'] or y['pid'] in used or y['pid'] == x['pid']:
                continue
            yn = set(n for n in (y['first'], y['nick']) if n)
            same_name = bool(xn & yn)
            alike = same_name or any((len(m) >= 3 and len(n) >= 3 and m[:3] == n[:3]) or (len(m) >= 4 and _lev(m, n) <= 1) for m in xn for n in yn)
            has_x, has_y = bool(x['bm'] and x['bd']), bool(y['bm'] and y['bd'])
            same_bd = has_x and has_y and x['bm'] == y['bm'] and x['bd'] == y['bd'] and (not x['by'] or not y['by'] or x['by'] == y['by'])
            close_bd = has_x and has_y and x['by'] and x['by'] == y['by'] and len(set([x['bm'], x['bd']]) & set([y['bm'], y['bd']])) > 0
            kind = 'same' if alike and same_bd else ('maybe' if same_name and (close_bd or not has_x or not has_y) else '')
            if kind:
                pairs.append([x['pid'], y['pid'], kind])
                used.add(y['pid'])
                break
    return pairs


def spell_counts(keep, lose):
    """How many other people use each spelling of the two first and last names, so a typo can
    lose to the usual spelling (Westfall over Westfll, Lindsey over Lindsy)."""
    out = {'first': {}, 'last': {}}
    rows = list(q.QuerySql("SELECT ISNULL(FirstName, '') F, ISNULL(LastName, '') L FROM dbo.People WITH (NOLOCK) WHERE PeopleId IN (%d, %d)" % (keep, lose)))
    for key, col, field in (('first', 'F', 'FirstName'), ('last', 'L', 'LastName')):
        vals = set(s(getattr(r, col)).strip() for r in rows)
        vals = [v for v in vals if v]
        if len(set(v.lower() for v in vals)) < 2:
            continue
        for v in vals:
            n = q.QuerySqlTop1("SELECT COUNT(*) N FROM dbo.People WITH (NOLOCK) WHERE %s = '%s' AND PeopleId NOT IN (%d, %d)" % (field, v.replace("'", "''"), keep, lose))
            out[key][v.lower()] = to_int(n.N) if n else 0
    return out


def older_family(k, others):
    """A child whose other record is in a family where every adult is 45 or more years older:
    grandparents who brought them, so moving the child in is not assumed."""
    ka = k.get('age')
    if ka is None or to_int(ka, 99) >= 18:
        return False
    adults = [to_int(x['age']) for x in others if x['age'] is not None and to_int(x['age']) >= 18]
    return bool(adults) and all(a - to_int(ka) >= 45 for a in adults)


def tag_plan(keep, lose):
    """Tags TouchPoint's merge would drop: lose's ordinary tags (status flags move),
    and the tags lose owns, which TouchPoint deletes with everyone in them."""
    member_of = [{'id': r.Id, 'name': s(r.Name), 'owner': r.Owner, 'ownerName': s(r.OwnerName), 'has': bool(r.Has)} for r in q.QuerySql("""
        SELECT t.Id, t.Name, ISNULL(t.PeopleId, 0) Owner, ISNULL(o.Name, '') OwnerName,
          CASE WHEN EXISTS (SELECT 1 FROM dbo.TagPerson k WITH (NOLOCK) WHERE k.Id = t.Id AND k.PeopleId = %(k)d) THEN 1 ELSE 0 END Has
        FROM dbo.TagPerson tp WITH (NOLOCK) JOIN dbo.Tag t WITH (NOLOCK) ON t.Id = tp.Id
        LEFT JOIN dbo.People o WITH (NOLOCK) ON o.PeopleId = t.PeopleId
        WHERE tp.PeopleId = %(l)d AND t.TypeId = 1 AND t.StatusFlagId IS NULL AND ISNULL(t.PeopleId, 0) <> %(l)d
        ORDER BY t.Name""" % {'k': keep, 'l': lose})]
    owned = [{'id': r.Id, 'name': s(r.Name), 'n': to_int(r.N)} for r in q.QuerySql("""
        SELECT t.Id, t.Name, (SELECT COUNT(*) FROM dbo.TagPerson tp WITH (NOLOCK) WHERE tp.Id = t.Id) N
        FROM dbo.Tag t WITH (NOLOCK) WHERE t.PeopleId = %d AND t.TypeId = 1 AND t.StatusFlagId IS NULL ORDER BY t.Name""" % lose)]
    return member_of, owned


def family_values(fid):
    out = []
    for r in q.QuerySql("""SELECT Field, ISNULL(Type, '') Type, ISNULL(StrValue, '') StrValue, ISNULL(Data, '') Data,
            CONVERT(varchar(10), DateValue, 120) DateValue, IntValue, BitValue
            FROM dbo.FamilyExtra WITH (NOLOCK) WHERE FamilyId = %d ORDER BY Field""" % fid):
        out.append({'field': s(r.Field), 'type': s(r.Type), 'str': s(r.StrValue), 'data': s(r.Data),
                    'date': s(r.DateValue), 'int': r.IntValue, 'bit': r.BitValue})
    return out


def nick_pair(keep, lose):
    """Two first names where one is a short form of the other (Judith and Judy, Pamela and Pam, William
    and Bill), and neither record has a Goes By yet. Uses TouchPoint's own nickname list, or the short
    one starting the long one. Returns {formal, short, formalSide} or None."""
    d = {}
    for x in q.QuerySql("SELECT PeopleId, ISNULL(FirstName, '') F, ISNULL(NickName, '') N FROM dbo.People WITH (NOLOCK) WHERE PeopleId IN (%d, %d)" % (keep, lose)):
        d[x.PeopleId] = (s(x.F).strip(), s(x.N).strip())
    if keep not in d or lose not in d:
        return None
    (kf, kn), (lf, ln) = d[keep], d[lose]
    if not kf or not lf or kf.lower() == lf.lower() or kn or ln:
        return None
    a, b = kf.lower(), lf.lower()
    qa, qb = a.replace("'", "''"), b.replace("'", "''")
    hit = q.QuerySqlTop1("""SELECT TOP 1 LOWER(FirstName) A FROM dbo.FirstNick WITH (NOLOCK) WHERE [count] >= 2
        AND ((LOWER(FirstName) = '%s' AND LOWER(NickName) = '%s') OR (LOWER(FirstName) = '%s' AND LOWER(NickName) = '%s'))""" % (qa, qb, qb, qa))
    if hit:
        formal = s(hit.A)
    else:
        short_, long_ = (a, b) if len(a) < len(b) else (b, a)
        if len(short_) >= 3 and long_.startswith(short_) and len(long_) - len(short_) >= 2:
            formal = long_
        else:
            return None
    if formal == a:
        return {'formal': kf, 'short': lf, 'formalSide': 'keep'}
    return {'formal': lf, 'short': kf, 'formalSide': 'lose'}


def carry_plan(keep, lose):
    """Person fields TouchPoint's merge does not carry over: the check-in note (printed and shown at
    children's check-in: allergies, custody) and the MinistrySafe account link (People.MSUserId; without
    it TouchPoint makes a second MinistrySafe account and the training history stays on the old one)."""
    d = {}
    for r in q.QuerySql("SELECT PeopleId, ISNULL(CheckInNotes, '') N, ISNULL(MSUserId, 0) M FROM dbo.People WITH (NOLOCK) WHERE PeopleId IN (%d, %d)" % (keep, lose)):
        d[r.PeopleId] = (s(r.N).strip(), to_int(r.M))
    if keep not in d or lose not in d:
        return {}
    (kn, km), (ln, lm) = d[keep], d[lose]
    out = {}
    if ln and ln.lower() not in kn.lower():
        out['checkInNote'] = (kn + ' | ' + ln) if kn else ln
        out['checkInBoth'] = bool(kn)
    if lm and not km:
        out['msUserId'] = lm
    return out


def email_fails(keep, lose):
    """Failed sends to the email addresses on either record (TouchPoint's EmailQueueToFail is kept by
    address): a bounce, a send dropped because the address failed before, or a spam report."""
    addrs = set()
    for r in q.QuerySql("SELECT LOWER(ISNULL(EmailAddress, '')) A, LOWER(ISNULL(EmailAddress2, '')) B FROM dbo.People WITH (NOLOCK) WHERE PeopleId IN (%d, %d)" % (keep, lose)):
        for a in (s(r.A).strip(), s(r.B).strip()):
            if '@' in a:
                addrs.add(a)
    out = {}
    if not addrs:
        return out
    try:
        rows = q.QuerySql("""SELECT LOWER(email) E, event Ev, ISNULL(bouncetype, '') B, CONVERT(varchar(10), MAX(time), 120) D,
              DATEDIFF(day, MAX(time), GETDATE()) Days, COUNT(*) N
            FROM dbo.EmailQueueToFail WITH (NOLOCK) WHERE email IN (%s) AND event IN ('bounce', 'dropped', 'spamreport')
            GROUP BY LOWER(email), event, ISNULL(bouncetype, '')""" % ','.join("'%s'" % a.replace("'", "''") for a in addrs))
    except Exception:
        return out
    for r in rows:
        e, ev, b, d = s(r.E), s(r.Ev), s(r.B), s(r.D)
        text = {'bounce': 'bounced' + (' (' + b.lower() + ')' if b else ''), 'dropped': 'not delivered (blocked after earlier failures)',
                'spamreport': 'marked a church email as spam'}.get(ev, ev)
        cur = out.get(e)
        if not cur or d > cur['last']:
            out[e] = {'text': text, 'last': d, 'recent': to_int(r.Days, 9999) <= 365, 'n': to_int(r.N) + (cur['n'] if cur else 0)}
        else:
            cur['n'] += to_int(r.N)
    return out


def other_losses(lose):
    """What TouchPoint's delete wipes and its merge never moves, counted for this record. Each check is
    on its own, so a table another church's database lacks is skipped."""
    checks = [
        ("SELECT COUNT(*) N FROM dbo.PeopleAuthorizedCheckOut WITH (NOLOCK) WHERE AuthorizedPeopleId = %(l)d",
         'child', 'children', ' #%(l)d was allowed to pick up at check-in. Add the person kept on those children again'),
        ("SELECT COUNT(*) N FROM dbo.PeopleCanEmailFor WITH (NOLOCK) WHERE OnBehalfOf = %(l)d OR CanEmail = %(l)d",
         '"can email for" link', '"can email for" links', ' (who gets whose email)'),
        ("SELECT (SELECT COUNT(*) FROM dbo.PeopleDirectorySettings WITH (NOLOCK) WHERE PeopleId = %(l)d) + (SELECT COUNT(*) FROM dbo.OrgPeopleDirectorySettings WITH (NOLOCK) WHERE PeopleId = %(l)d) N",
         'directory privacy choice', 'directory privacy choices', ': set them again on the person kept'),
        ("SELECT COUNT(*) N FROM dbo.PeopleSignatureRequests WITH (NOLOCK) WHERE PeopleId = %(l)d", 'e-signature request', 'e-signature requests', ''),
        ("SELECT COUNT(*) N FROM dbo.ProcessPeople WITH (NOLOCK) WHERE PeopleId = %(l)d", 'Process Builder process in progress', 'Process Builder processes in progress', ''),
        ("SELECT COUNT(*) N FROM dbo.Coupons WITH (NOLOCK) WHERE PeopleId = %(l)d", 'registration coupon', 'registration coupons', ''),
        ("SELECT COUNT(*) N FROM dbo.VolInterestInterestCodes WITH (NOLOCK) WHERE PeopleId = %(l)d", 'volunteer interest', 'volunteer interests', ''),
        ("SELECT COUNT(*) N FROM dbo.SubRequest WITH (NOLOCK) WHERE RequestorId = %(l)d OR SubstituteId = %(l)d", 'substitute request', 'substitute requests', ''),
        ("SELECT COUNT(*) N FROM dbo.EmailResponses WITH (NOLOCK) WHERE PeopleId = %(l)d", 'email open', 'email opens', ' (the record of which church emails #%(l)d opened)'),
    ]
    out = []
    for sql, one, many, tail in checks:
        try:
            r = q.QuerySqlTop1(sql % {'l': lose})
            n = to_int(r.N) if r else 0
        except Exception:
            n = 0
        if n:
            out.append('%d %s%s' % (n, one if n == 1 else many, tail % {'l': lose}))
    return out


def act_pair():
    keep, lose = pair_ids()
    k, l = person_detail(keep), person_detail(lose)
    if not k or not l:
        return {'success': False, 'message': 'One of these people is no longer in TouchPoint (already merged?).'}
    same_family = k['fid'] == l['fid']
    kf = family_of(k['fid'], (keep, lose))
    lf = kf if same_family else family_of(l['fid'], (keep, lose))
    member_of, owned = tag_plan(keep, lose)
    lose_vals = family_values(l['fid'])
    keep_fields = set(v['field'] for v in (lose_vals if same_family else family_values(k['fid'])))
    empty_after = l['familySize'] <= 1 and not same_family
    # TouchPoint deletes the deleted record's family extra values even when that family still has
    # people in it: put them back on that same family; copy them over only when the family is gone
    others = [x for x in lf if not x['self']]
    mode = 'same' if same_family else ('their' if others else 'copy')
    loss = {
        'tags': member_of, 'ownedTags': owned,
        'familyValues': [{'field': v['field'], 'kept': mode == 'copy' and v['field'] in keep_fields} for v in lose_vals],
        'valuesMode': mode,
        'sameFamily': same_family,
        'familyGone': empty_after,
        'relatedGone': len(related_of(l['fid'])) if empty_after else 0,
        'headless': l['isHead'] and l['familySize'] > 1,
        # the deleted record is the one living with others: merged this way, the person kept is not in that family
        'leavesFamily': [] if same_family else [{'pid': x['pid'], 'name': x['name'], 'age': x['age']} for x in others],
        'keepOthers': [] if same_family else [{'pid': x['pid'], 'name': x['name'], 'age': x['age']} for x in kf if not x['self']],
        'keepAlone': (not same_family) and k['familySize'] <= 1,
        'olderFamily': (not same_family) and older_family(k, others),
        'carry': carry_plan(keep, lose), 'more': other_losses(lose),
        'absent': l['counts']['absent'], 'cards': l['counts']['cards'],
    }
    return {'success': True, 'keep': k, 'lose': l, 'keepFamily': kf, 'loseFamily': lf, 'sameFamily': same_family, 'nick': nick_pair(keep, lose), 'spell': spell_counts(keep, lose), 'emailFails': email_fails(keep, lose),
            'familyMatches': [] if same_family else match_families(kf, lf),
            'keepRelated': related_of(k['fid']), 'loseRelated': [] if same_family else related_of(l['fid']),
            'loss': loss, 'values': value_plan(keep, lose), 'medical': medical_plan(keep, lose)}


# ---------------------------------------------------------------- person extra values and medical
SETTABLE = ('Code', 'Text', 'Date', 'Int', 'Bit')


def person_values(pid):
    """A person's extra values, one per field, with the value as text for showing."""
    out = {}
    for r in q.QuerySql("""SELECT Field, ISNULL(Type, '') Type, ISNULL(StrValue, '') StrValue, ISNULL(Data, '') Data,
            CONVERT(varchar(10), DateValue, 120) DateValue, IntValue, BitValue
            FROM dbo.PeopleExtra WITH (NOLOCK) WHERE PeopleId = %d ORDER BY Field""" % pid):
        f, ty = s(r.Field), s(r.Type)
        if ty == 'Code':
            v = s(r.StrValue)
        elif ty == 'Text':
            v = s(r.Data)
        elif ty == 'Date':
            v = s(r.DateValue)
        elif ty == 'Int':
            v = '' if r.IntValue is None else unicode(r.IntValue)
        elif ty == 'Bit':
            v = '' if r.BitValue is None else ('Yes' if r.BitValue else 'No')
        else:
            v = s(r.StrValue) or s(r.Data) or s(r.DateValue) or ('' if r.IntValue is None else unicode(r.IntValue))
        if f not in out:
            out[f] = {'field': f, 'type': ty, 'value': v[:300]}
    return out


# ---------------------------------------------------------------- the merge log
# One log per person kept: a person extra value, registered in StandardExtraValues2 with
# VisibilityRoles so only Admin and Manager2 see it on a profile (the TPxi Benevolence way).
# An Admin opening the page registers it (and the Not the same marks) when needed; until it is
# registered nothing is logged. Never Special Content: TouchPoint serves that with no sign-in.
# It never keeps giving: no gift counts, statements, envelopes or cards.
LOG_FIELD = 'TPxiDuplicatesLog'
LOG_NAMES = [LOG_FIELD, LOG_FIELD + '_mv', LOG_FIELD + '_mv_mv']          # a merge renames a second copy
OWN_FIELDS = LOG_NAMES + [NOT_SAME, NOT_SAME + '_mv', NOT_SAME + '_mv_mv']
SEV_CONTENT = 'StandardExtraValues2'
GIVING = re.compile(r'gift|giving|contribut|envelope|statement|pledge|\bfunds?\b|card|bank|routing', re.I)


def sev_xml():
    try:
        return unicode(model.TextContent(SEV_CONTENT) or '')
    except Exception:
        return ''


def log_protected():
    """Every name we keep is registered as seen only by the tool's roles."""
    xml = sev_xml()
    allowed = [r.lower() for r in ROLES]
    for name in OWN_FIELDS:
        m = re.search(r'<Value\s+Name="%s"[^>]*>' % re.escape(name), xml)
        roles = re.search(r'VisibilityRoles="([^"]*)"', m.group(0)) if m else None
        have = [x.strip().lower() for x in roles.group(1).split(',') if x.strip()] if roles else []
        if not have or any(h not in allowed for h in have):
            return False
    return True


def protect_now():
    """Admin: register (or fix) our extra value names in StandardExtraValues2 as seen only by
    the tool's roles, leaving every other entry as it was."""
    xml = sev_xml() or '<?xml version="1.0" encoding="utf-16"?>\n<Views>\n</Views>'
    for name in OWN_FIELDS:
        line = '<Value Name="%s" Type="Text" VisibilityRoles="%s" />' % (name, ','.join(ROLES))
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


def me():
    try:
        pid = to_int(model.UserPeopleId)   # a property, not a method
    except Exception:
        pid = 0
    r = q.QuerySqlTop1('SELECT Name FROM dbo.People WITH (NOLOCK) WHERE PeopleId = %d' % pid) if pid else None
    return pid, (s(r.Name) if r else '')


def now_text():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M')


def log_read(pid):
    """All entries on a person, from the log and any copies a merge renamed, oldest first."""
    entries, names = [], []
    for r in q.QuerySql("SELECT Field, ISNULL(Data, '') D FROM dbo.PeopleExtra WITH (NOLOCK) WHERE PeopleId = %d AND Field IN (%s)" % (
            pid, ','.join("'%s'" % n for n in LOG_NAMES))):
        names.append(s(r.Field))
        entries += load_json(s(r.D), {}).get('entries') or []
    seen, out = set(), []
    for e in sorted(entries, key=lambda e: e.get('t') or ''):
        if e.get('id') in seen:
            continue
        seen.add(e.get('id'))
        out.append(e)
    return out, names


def log_write(pid, entries, names):
    who = 'PeopleId = %d' % pid
    model.AddExtraValueText(who, LOG_FIELD, safe_json({'v': 1, 'entries': entries[-40:]}))
    for n in names:
        if n != LOG_FIELD:
            model.DeleteExtraValue(who, n)


def log_add(pid, entry):
    entries, names = log_read(pid)
    entries.append(entry)
    log_write(pid, entries, names)


def snapshot(pid):
    """What the record to be deleted holds that the merge does not keep anywhere else: who it
    was, how to reach them, and their family. Nothing about giving."""
    r = q.QuerySqlTop1("""SELECT p.Name, ISNULL(p.FirstName, '') F, ISNULL(p.NickName, '') Nk, ISNULL(p.MiddleName, '') Mi, ISNULL(p.LastName, '') L,
          ISNULL(p.MaidenName, '') Mn, p.BirthMonth BM, p.BirthDay BD, p.BirthYear BYr, ISNULL(g.Description, '') G, ISNULL(mar.Description, '') Mar,
          ISNULL(ms.Description, '') MS, ISNULL(p.EmailAddress, '') E, ISNULL(p.EmailAddress2, '') E2, ISNULL(p.CellPhone, '') C, ISNULL(p.WorkPhone, '') W,
          ISNULL(f.HomePhone, '') H, ISNULL(f.AddressLineOne, '') A1, ISNULL(f.AddressLineTwo, '') A2, ISNULL(f.CityName, '') City,
          ISNULL(f.StateCode, '') St, ISNULL(f.ZipCode, '') Zip, CONVERT(varchar(10), p.CreatedDate, 120) Added,
          CONVERT(varchar(10), p.JoinDate, 120) Joined, CONVERT(varchar(10), p.BaptismDate, 120) Baptized,
          ISNULL(o.Description, '') O, ISNULL(ep.Description, '') EP, p.FamilyId Fam
        FROM dbo.People p WITH (NOLOCK) JOIN dbo.Families f WITH (NOLOCK) ON f.FamilyId = p.FamilyId
        LEFT JOIN lookup.Gender g ON g.Id = p.GenderId LEFT JOIN lookup.MaritalStatus mar ON mar.Id = p.MaritalStatusId
        LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId LEFT JOIN lookup.Origin o ON o.Id = p.OriginId
        LEFT JOIN lookup.EntryPoint ep ON ep.Id = p.EntryPointId
        WHERE p.PeopleId = %d""" % pid)
    if not r:
        return {'pid': pid}
    bd = '/'.join(unicode(x) for x in (r.BM, r.BD, r.BYr) if x)
    addr = ', '.join(x for x in (s(r.A1), s(r.A2), ' '.join(y for y in (s(r.City), s(r.St), s(r.Zip)) if y)) if x)
    fam = ['%s%s%s' % (x['name'], ', %s' % x['age'] if x['age'] is not None else '', ', ' + x['position'] if x['position'] else '')
           for x in family_of(to_int(r.Fam), (pid,)) if not x['self']]
    d = {'pid': pid, 'name': s(r.Name), 'first': s(r.F), 'nick': s(r.Nk), 'middle': s(r.Mi), 'last': s(r.L), 'former': s(r.Mn),
         'birthday': bd, 'gender': s(r.G), 'marital': s(r.Mar), 'member': s(r.MS), 'email': s(r.E), 'email2': s(r.E2),
         'cell': s(r.C), 'work': s(r.W), 'home': s(r.H), 'address': addr, 'added': s(r.Added), 'joined': s(r.Joined),
         'baptized': s(r.Baptized), 'how': clean_how(r.O, r.EP), 'family': fam}
    return dict((k, v) for k, v in d.items() if v)


def act_log():
    """After the merge: what happened, step by step, onto the entry the before step started."""
    keep, eid = to_int(_param('keep')), _param('id')
    if not keep or not eid or not log_protected():
        return {'success': True}
    steps, problems = [], []
    for x in load_json(_param('lines'), [])[:80]:
        text = unicode((x or {}).get('text') or '')[:300]
        if not text or GIVING.search(text):
            continue
        (problems if x.get('cls') == 'bad' else steps).append(text)
    entries, names = log_read(keep)
    for e in entries:
        if e.get('id') == eid:
            e['status'] = 'check' if problems else 'done'
            e['steps'], e['problems'] = steps, problems
    log_write(keep, entries, names)
    return {'success': True}


def act_merges():
    """The newest entries across everyone, for the Merge log on the list."""
    if not log_protected():
        return {'success': True, 'on': False, 'entries': []}
    out, seen = [], set()
    for r in q.QuerySql("""SELECT TOP 500 e.PeopleId, p.Name, ISNULL(e.Data, '') D FROM dbo.PeopleExtra e WITH (NOLOCK)
            JOIN dbo.People p WITH (NOLOCK) ON p.PeopleId = e.PeopleId WHERE e.Field IN (%s)""" % ','.join("'%s'" % n for n in LOG_NAMES)):
        for e in load_json(s(r.D), {}).get('entries') or []:
            if e.get('id') in seen:
                continue
            seen.add(e.get('id'))
            e['at'] = {'pid': r.PeopleId, 'name': s(r.Name)}
            out.append(e)
    out.sort(key=lambda e: e.get('t') or '', reverse=True)
    return {'success': True, 'on': True, 'entries': out[:40]}


def value_plan(keep, lose):
    """TouchPoint moves the other record's extra values, but a field the kept person already
    has comes over renamed Field_mv. Same value on both: the copy is just clutter. Different
    values: someone should choose."""
    # this tool's own values (the merge log, Not the same) are settled by it, never offered as a choice
    kv = dict((k, v) for k, v in person_values(keep).items() if k not in OWN_FIELDS)
    lv = dict((k, v) for k, v in person_values(lose).items() if k not in OWN_FIELDS)
    conflicts, dups, moves = [], [], []
    for f, o in sorted(lv.items()):
        if f == NOT_SAME:
            continue
        if f in kv:
            if kv[f]['value'].strip().lower() == o['value'].strip().lower():
                dups.append(f)
            else:
                conflicts.append({'field': f, 'type': o['type'], 'keep': kv[f]['value'], 'other': o['value'],
                                  'settable': o['type'] in SETTABLE and kv[f]['type'] == o['type']})
        else:
            moves.append(f)
    return {'keep': sorted(kv.values(), key=lambda x: x['field'].lower()), 'lose': sorted(lv.values(), key=lambda x: x['field'].lower()),
            'conflicts': conflicts, 'dups': dups, 'moves': moves}


MED = [('MedicalDescription', 'Allergies and medical'), ('Emcontact', 'Emergency contact'), ('Emphone', 'Emergency phone'),
       ('Doctor', 'Doctor'), ('Docphone', 'Doctor phone'), ('Insurance', 'Insurance'), ('Policy', 'Policy'), ('ShirtSize', 'Shirt size')]


def medical(pid):
    r = q.QuerySqlTop1('SELECT ' + ', '.join("ISNULL(%s, '') %s" % (c, c) for c, _ in MED) + ' FROM dbo.RecReg WITH (NOLOCK) WHERE PeopleId = %d' % pid)
    return dict((c, s(getattr(r, c)) if r else '') for c, _ in MED)


def medical_plan(keep, lose):
    """When both records have medical and emergency info, TouchPoint's merge takes the other
    record's value for each of these wherever it has one, over the kept person's."""
    km, lm = medical(keep), medical(lose)
    rows = []
    for c, label in MED:
        a, b = km[c], lm[c]
        if not a and not b:
            continue
        what = 'same' if a.lower() == b.lower() else ('replaced' if a and b else ('filled' if b else 'kept'))
        rows.append({'field': c, 'label': label, 'keep': a, 'other': b, 'what': what})
    return rows


# ---------------------------------------------------------------- around the merge
def union_counts(keep, lose):
    """What the person kept should hold once both records are one."""
    r = q.QuerySqlTop1(gift_sql("""
        SELECT
          GIFTS_KL Gifts,
          (SELECT COUNT(DISTINCT MeetingId) FROM dbo.Attend WITH (NOLOCK) WHERE PeopleId IN (%(k)d, %(l)d) AND AttendanceFlag = 1) Att,
          (SELECT COUNT(DISTINCT OrganizationId) FROM dbo.OrganizationMembers WITH (NOLOCK) WHERE PeopleId IN (%(k)d, %(l)d)) Orgs,
          (SELECT COUNT(*) FROM dbo.Users WITH (NOLOCK) WHERE PeopleId IN (%(k)d, %(l)d)) Logins,
          (SELECT COUNT(*) FROM dbo.BackgroundChecks WITH (NOLOCK) WHERE PeopleID IN (%(k)d, %(l)d)) Checks,
          (SELECT COUNT(*) FROM dbo.TaskNote WITH (NOLOCK) WHERE AboutPersonId IN (%(k)d, %(l)d)) Notes""") % {'k': keep, 'l': lose})
    return {'gifts': to_int(r.Gifts), 'att': to_int(r.Att), 'orgs': to_int(r.Orgs), 'logins': to_int(r.Logins),
            'checks': to_int(r.Checks), 'notes': to_int(r.Notes)}


def act_before():
    """Right before TouchPoint's merge: put the person kept into the tags TouchPoint
    would drop, rebuild the tags the other record owns (now owned by the person
    kept), and hand back the other family's extra values and what to expect after."""
    keep, lose = pair_ids()
    member_of, owned = tag_plan(keep, lose)
    log = []
    for t in member_of:
        if t['has'] or not t['owner']:
            continue
        model.AddTag('PeopleId = %d' % keep, t['name'], t['owner'])
        log.append('Added to the tag %s (%s)' % (t['name'], t['ownerName'] or 'owner %d' % t['owner']))
    for t in owned:
        ids = [r.PeopleId for r in q.QuerySql('SELECT PeopleId FROM dbo.TagPerson WITH (NOLOCK) WHERE Id = %d' % t['id'])]
        ids = [keep if i == lose else i for i in ids]
        ids = sorted(set(ids))
        if ids:
            model.AddTag('PeopleId IN (%s)' % ','.join(str(i) for i in ids), t['name'], keep)
            log.append('Rebuilt the tag %s (%d people), now owned by the person kept' % (t['name'], len(ids)))
    lp = q.QuerySqlTop1('SELECT FamilyId F FROM dbo.People WITH (NOLOCK) WHERE PeopleId = %d' % lose)
    fam_snap = family_values(lp.F) if lp else []
    # where the values go back: anyone else still in that family (its own family), else the person kept
    anchor = 0
    if lp:
        a = q.QuerySqlTop1('SELECT TOP 1 PeopleId P FROM dbo.People WITH (NOLOCK) WHERE FamilyId = %d AND PeopleId NOT IN (%d, %d) ORDER BY PositionInFamilyId, PeopleId' % (lp.F, keep, lose))
        anchor = a.P if a else 0
    keep_fields = [s(r.Field) for r in q.QuerySql('SELECT Field FROM dbo.PeopleExtra WITH (NOLOCK) WHERE PeopleId = %d' % keep)]
    # the merge log: started now, with what the record to be deleted holds, so it survives the delete
    log_id = ''
    if log_protected():
        pid, nm = me()
        kn = q.QuerySqlTop1('SELECT Name FROM dbo.People WITH (NOLOCK) WHERE PeopleId = %d' % keep)
        log_id = 'm%s-%d' % (datetime.datetime.now().strftime('%Y%m%d%H%M%S'), lose)
        picks = [x for x in load_json(_param('picks'), []) if isinstance(x, dict) and not GIVING.search(unicode(x.get('field') or ''))][:60]
        try:
            log_add(keep, {'id': log_id, 't': now_text(), 'kind': 'merge', 'status': 'started', 'by': nm, 'byId': pid, 'v': VERSION,
                           'into': {'pid': keep, 'name': s(kn.Name) if kn else ''}, 'from': snapshot(lose), 'picks': picks,
                           'move': _param('move')[:300]})
        except Exception as e:
            log_id = ''
            log.append('The merge log could not be started: %s' % unicode(e)[:120])
    return {'success': True, 'log': log, 'familyValues': fam_snap, 'anchor': anchor, 'expect': union_counts(keep, lose), 'keepFields': keep_fields,
            'carry': carry_plan(keep, lose), 'logId': log_id}


def act_after():
    """After TouchPoint's merge: give the kept person's family any extra values it is
    missing from the snapshot (TouchPoint deletes the other record's family values,
    which is the kept family's own when both were in one family)."""
    keep = to_int(_param('keep'))
    target = to_int(_param('anchor')) or keep   # the deleted record's family if people remain in it
    snap = load_json(_param('familyValues'), [])
    log = []
    # the person fields TouchPoint's merge leaves behind (snapshot taken before the merge)
    carry = load_json(_param('carry'), {})
    if carry:
        person = model.GetPerson(keep)
        if carry.get('checkInNote'):
            try:
                model.UpdateField(person, 'CheckInNotes', carry['checkInNote'])
                log.append('Check-in note carried over' + (' (both notes kept, joined with |)' if carry.get('checkInBoth') else ''))
            except Exception as e:
                log.append('Check-in note could not be carried over: %s' % unicode(e)[:120])
        if carry.get('msUserId'):
            try:
                model.UpdateField(person, 'MSUserId', to_int(carry['msUserId']))
                log.append('MinistrySafe account link kept, so MinistrySafe training and checks stay with this person')
            except Exception as e:
                log.append('MinistrySafe account link could not be kept: %s' % unicode(e)[:120])
    tf = q.QuerySqlTop1('SELECT FamilyId F FROM dbo.People WITH (NOLOCK) WHERE PeopleId = %d' % target)
    if not tf:
        return {'success': True, 'log': log}
    have = set(v['field'] for v in family_values(to_int(tf.F)))
    who = 'PeopleId = %d' % target
    for v in snap[:200]:
        f = unicode(v.get('field') or '')
        if not f or f in have:
            continue
        t = (v.get('type') or '').lower()
        try:
            if t == 'code':
                model.AddExtraValueCodeFamily(who, f, v.get('str') or '')
            elif t == 'date' and v.get('date'):
                model.AddExtraValueDateFamily(who, f, v.get('date'))
            elif t == 'int' and v.get('int') is not None:
                model.AddExtraValueIntFamily(who, f, to_int(v.get('int')))
            elif t == 'bit':
                model.AddExtraValueBoolFamily(who, f, bool(v.get('bit')))
            else:
                model.AddExtraValueTextFamily(who, f, v.get('data') or v.get('str') or '')
            log.append('Family value %s put back' % f)
        except Exception as e:
            log.append('Family value %s could not be put back: %s' % (f, unicode(e)[:120]))
    return {'success': True, 'log': log}


def act_values():
    """After TouchPoint's merge: settle the extra values that came over as Field_mv.
    picks: {field: 'keep' | 'other'}; dups: fields that had the same value on both."""
    keep = to_int(_param('keep'))
    picks = load_json(_param('picks'), {})
    dups = load_json(_param('dups'), [])
    before = set(load_json(_param('keepFields'), []))
    now = person_values(keep)
    raw = dict((s(r.Field), r) for r in q.QuerySql("""SELECT Field, ISNULL(Type, '') Type, ISNULL(StrValue, '') StrValue, ISNULL(Data, '') Data,
        CONVERT(varchar(10), DateValue, 120) DateValue, IntValue, BitValue FROM dbo.PeopleExtra WITH (NOLOCK) WHERE PeopleId = %d""" % keep))
    who = 'PeopleId = %d' % keep
    log = []

    def moved(f):
        # the copy TouchPoint just made: Field_mv (or _mv_mv...), new since before the merge
        c = [x for x in now if x.startswith(f + '_mv') and x not in before]
        return sorted(c, key=len)[0] if c else None

    for f in list(dups)[:200]:
        m = moved(f)
        if m:
            model.DeleteExtraValue(who, m)
            log.append('Extra value %s: the same on both, so the extra copy (%s) was removed' % (f, m))
    for f, pick in list(picks.items())[:200]:
        m = moved(f)
        if not m:
            continue
        r = raw.get(m)
        if pick == 'other' and r is not None:
            ty = s(r.Type)
            try:
                if ty == 'Code':
                    model.AddExtraValueCode(who, f, s(r.StrValue))
                elif ty == 'Text':
                    model.AddExtraValueText(who, f, s(r.Data))
                elif ty == 'Date':
                    model.AddExtraValueDate(who, f, s(r.DateValue))
                elif ty == 'Int':
                    model.AddExtraValueInt(who, f, to_int(r.IntValue))
                elif ty == 'Bit':
                    model.AddExtraValueBool(who, f, bool(r.BitValue))
                else:
                    log.append('Extra value %s: could not be set (type %s); both are kept, the other as %s' % (f, ty or 'unknown', m))
                    continue
            except Exception as e:
                log.append('Extra value %s: could not be set (%s); both are kept, the other as %s' % (f, unicode(e)[:100], m))
                continue
            model.DeleteExtraValue(who, m)
            log.append('Extra value %s: now the other record\'s value' % f)
        elif pick == 'keep':
            model.DeleteExtraValue(who, m)
            log.append('Extra value %s: kept as it was; the other record\'s value was removed' % f)
    # Not the same marks from both records, as one list
    m = moved(NOT_SAME)
    if m:
        ids = set(i for i in (s(getattr(raw.get(NOT_SAME), 'Data', '')) + ',' + s(raw[m].Data)).split(',') if i.strip())
        ids.discard(str(keep))
        model.AddExtraValueText(who, NOT_SAME, ',' + ','.join(sorted(ids)) + ',')
        model.DeleteExtraValue(who, m)
    return {'success': True, 'log': log}


def act_verify():
    keep, lose = pair_ids()
    expect = load_json(_param('expect'), {})
    gone = q.QuerySqlTop1('SELECT COUNT(*) N FROM dbo.People WITH (NOLOCK) WHERE PeopleId = %d' % lose).N == 0
    now = union_counts(keep, 0)
    names = [('gifts', 'gift', 'gifts'), ('att', 'meeting attended', 'meetings attended'), ('orgs', 'involvement', 'involvements'),
             ('logins', 'login', 'logins'), ('checks', 'background check', 'background checks'), ('notes', 'note or task', 'notes and tasks')]
    items = []
    for key, one, label in names:
        if key == 'gifts' and not FINANCE:
            continue   # not readable without the Finance role; TouchPoint's merge moves every gift
        want, got = to_int(expect.get(key)), now[key]
        items.append({'what': one if got == 1 else label, 'want': want, 'got': got, 'ok': got >= want})
    # the name now, which a merge can change (taking a married name, a typed former name)
    p = q.QuerySqlTop1('SELECT ISNULL(Name, \'\') N, ISNULL(MaidenName, \'\') M FROM dbo.People WITH (NOLOCK) WHERE PeopleId = %d' % keep)
    return {'success': True, 'gone': gone, 'items': items, 'name': s(p.N) if p else '', 'former': s(p.M) if p else ''}


def _photo(pid):
    r = q.QuerySqlTop1("""SELECT ISNULL(pic.MediumUrl, '') U, ISNULL(pic.LargeUrl, '') L, ISNULL(pic.MediumId, 0) M
        FROM dbo.People p WITH (NOLOCK) JOIN dbo.Picture pic WITH (NOLOCK) ON pic.PictureId = p.PictureId WHERE p.PeopleId = %d""" % pid)
    if not r:
        return ''
    return s(r.U) or s(r.L) or ('/api/v1/Portrait/%d' % to_int(r.M) if to_int(r.M) > 0 else '')


def _strip(html, n):
    t = re.sub(r'<[^>]+>', ' ', unicode(html or ''))
    t = re.sub(r'&nbsp;', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t[:n] + ('...' if len(t) > n else '')


def act_detail():
    """What is inside one row of What each holds, for both records, and what they share."""
    keep, lose = pair_ids()
    kind = _param('kind')
    out = {'success': True, 'kind': kind}
    if kind == 'photo':
        out['keep'], out['lose'] = _photo(keep), _photo(lose)
    elif kind == 'att':
        for side, pid in (('keep', keep), ('lose', lose)):
            out[side] = {
                'recent': [{'d': s(r.D), 'org': s(r.O)} for r in q.QuerySql("""
                    SELECT TOP 12 CONVERT(varchar(10), a.MeetingDate, 120) D, o.OrganizationName O FROM dbo.Attend a WITH (NOLOCK)
                    JOIN dbo.Organizations o WITH (NOLOCK) ON o.OrganizationId = a.OrganizationId
                    WHERE a.PeopleId = %d AND a.AttendanceFlag = 1 ORDER BY a.MeetingDate DESC""" % pid)],
                'top': [{'org': s(r.O), 'n': to_int(r.N), 'last': s(r.L)} for r in q.QuerySql("""
                    SELECT TOP 8 o.OrganizationName O, COUNT(*) N, CONVERT(varchar(10), MAX(a.MeetingDate), 120) L FROM dbo.Attend a WITH (NOLOCK)
                    JOIN dbo.Organizations o WITH (NOLOCK) ON o.OrganizationId = a.OrganizationId
                    WHERE a.PeopleId = %d AND a.AttendanceFlag = 1 GROUP BY o.OrganizationName ORDER BY COUNT(*) DESC""" % pid)]}
        out['same'] = [{'d': s(r.D), 'org': s(r.O)} for r in q.QuerySql("""
            SELECT TOP 10 CONVERT(varchar(10), a.MeetingDate, 120) D, o.OrganizationName O FROM dbo.Attend a WITH (NOLOCK)
            JOIN dbo.Attend b WITH (NOLOCK) ON b.MeetingId = a.MeetingId AND b.PeopleId = %d AND b.AttendanceFlag = 1
            JOIN dbo.Organizations o WITH (NOLOCK) ON o.OrganizationId = a.OrganizationId
            WHERE a.PeopleId = %d AND a.AttendanceFlag = 1 ORDER BY a.MeetingDate DESC""" % (lose, keep))]
        out['sameN'] = to_int(q.QuerySqlTop1("""SELECT COUNT(*) N FROM dbo.Attend a WITH (NOLOCK)
            JOIN dbo.Attend b WITH (NOLOCK) ON b.MeetingId = a.MeetingId AND b.PeopleId = %d AND b.AttendanceFlag = 1
            WHERE a.PeopleId = %d AND a.AttendanceFlag = 1""" % (lose, keep)).N)
    elif kind == 'orgs':
        for side, pid in (('keep', keep), ('lose', lose)):
            out[side] = [{'oid': r.OrganizationId, 'org': s(r.O), 'as': s(r.MT), 'since': s(r.D)} for r in q.QuerySql("""
                SELECT o.OrganizationId, o.OrganizationName O, ISNULL(mt.Description, '') MT, CONVERT(varchar(10), om.EnrollmentDate, 120) D
                FROM dbo.OrganizationMembers om WITH (NOLOCK) JOIN dbo.Organizations o WITH (NOLOCK) ON o.OrganizationId = om.OrganizationId
                LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId WHERE om.PeopleId = %d ORDER BY om.EnrollmentDate DESC""" % pid)]
    elif kind == 'notes':
        # only the notes this person may see in TouchPoint: no role limit, or one of their roles
        me = to_int(model.UserPeopleId)
        roles = [str(r.RoleId) for r in q.QuerySql('SELECT ur.RoleId FROM dbo.UserRole ur WITH (NOLOCK) JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = ur.UserId WHERE u.PeopleId = %d' % me)] if me else []
        allowed = '(t.RoleId IS NULL' + (' OR t.RoleId IN (%s)' % ','.join(roles) if roles else '') + ')'
        for side, pid in (('keep', keep), ('lose', lose)):
            rows = q.QuerySql("""
                SELECT TOP 15 CONVERT(varchar(10), t.CreatedDate, 120) D, ISNULL(t.IsNote, 0) IsNote, ISNULL(cp.Name, '') ByName,
                  ISNULL(t.Notes, '') N, ISNULL(t.Instructions, '') I
                FROM dbo.TaskNote t WITH (NOLOCK) LEFT JOIN dbo.People cp WITH (NOLOCK) ON cp.PeopleId = t.CreatedBy
                WHERE t.AboutPersonId = %d AND %s ORDER BY t.CreatedDate DESC""" % (pid, allowed))
            hidden = to_int(q.QuerySqlTop1('SELECT COUNT(*) N FROM dbo.TaskNote t WITH (NOLOCK) WHERE t.AboutPersonId = %d AND NOT %s' % (pid, allowed)).N)
            out[side] = {'items': [{'d': s(r.D), 'note': bool(r.IsNote), 'by': s(r.ByName), 'text': _strip(r.N or r.I, 220)} for r in rows], 'hidden': hidden}
    elif kind == 'tags':
        for side, pid in (('keep', keep), ('lose', lose)):
            out[side] = [{'id': r.Id, 'name': s(r.Name), 'owner': s(r.Owner), 'flag': bool(r.Flag)} for r in q.QuerySql("""
                SELECT t.Id, t.Name, ISNULL(o.Name, '') Owner, CASE WHEN t.StatusFlagId IS NOT NULL THEN 1 ELSE 0 END Flag
                FROM dbo.TagPerson tp WITH (NOLOCK) JOIN dbo.Tag t WITH (NOLOCK) ON t.Id = tp.Id LEFT JOIN dbo.People o WITH (NOLOCK) ON o.PeopleId = t.PeopleId
                WHERE tp.PeopleId = %d AND ((t.TypeId = 1 AND t.StatusFlagId IS NULL) OR t.StatusFlagId IS NOT NULL) ORDER BY Flag, t.Name""" % pid)]
    elif kind == 'timeline':
        out['keep'], out['lose'] = timeline(keep), timeline(lose)
    elif kind == 'login':
        for side, pid in (('keep', keep), ('lose', lose)):
            # TouchPoint mostly records use as last activity (set to the creation date when made), not last login
            out[side] = [{'user': s(r.Username), 'created': s(r.C), 'last': s(r.L), 'locked': bool(r.IsLockedOut)} for r in q.QuerySql("""
                SELECT Username, CONVERT(varchar(10), CreationDate, 120) C,
                  CONVERT(varchar(10), CASE WHEN ISNULL(LastLoginDate, '1900-01-01') > ISNULL(LastActivityDate, '1900-01-01') THEN LastLoginDate ELSE LastActivityDate END, 120) L,
                  ISNULL(IsLockedOut, 0) IsLockedOut
                FROM dbo.Users WITH (NOLOCK) WHERE PeopleId = %d""" % pid)]
    else:
        return {'success': False, 'message': 'Unknown detail.'}
    return out


def act_notsame():
    a, b = to_int(_param('a')), to_int(_param('b'))
    if not a or not b or a == b:
        return {'success': False, 'message': 'Pick two different people.'}
    for x, y in ((a, b), (b, a)):
        r = q.QuerySqlTop1("SELECT ISNULL(Data, '') D FROM dbo.PeopleExtra WITH (NOLOCK) WHERE PeopleId = %d AND Field = '%s'" % (x, NOT_SAME))
        have = [i for i in (r.D if r else '').split(',') if i.strip()]
        if str(y) not in have:
            have.append(str(y))
        model.AddExtraValueText('PeopleId = %d' % x, NOT_SAME, ',' + ','.join(have[-200:]) + ',')
    if log_protected():
        pid, nm = me()
        nmz = dict((r.PeopleId, s(r.Name)) for r in q.QuerySql('SELECT PeopleId, Name FROM dbo.People WITH (NOLOCK) WHERE PeopleId IN (%d, %d)' % (a, b)))
        eid = 'n%s-%d-%d' % (datetime.datetime.now().strftime('%Y%m%d%H%M%S'), min(a, b), max(a, b))
        for x, y in ((a, b), (b, a)):
            try:
                log_add(x, {'id': eid, 't': now_text(), 'kind': 'notsame', 'by': nm, 'byId': pid, 'v': VERSION,
                            'other': {'pid': y, 'name': nmz.get(y, '')}})
            except Exception:
                pass
    return {'success': True}


def dc_headers():
    """Sent with the update download: this church's TouchPoint address and the script and
    version, so the TPxi script service can say who is on what. No person, nothing from the
    database. The admin setting TPxiSendHost = 0 turns it off (shared with the other TPxi scripts)."""
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
    h["X-TPxi-Version"] = VERSION
    return h


def apply_update():
    """Admin only (checked in handle_post). Saves the newest TPxi Duplicates over this script,
    under whatever name it was saved as. Saved marks are extra values, so they are untouched."""
    name = re.sub(r'[^\w.\- ]', '', _param('script_name') or '') or DC_SCRIPT_ID
    try:
        code = model.RestGet(DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID, dc_headers()) or ''
    except Exception as e:
        return {'success': False, 'message': 'Could not download the update: %s' % e}
    code = unicode(code)
    # never save something that is not this script: an error page, an empty reply or
    # another tool would break the page for everyone
    if len(code) < 5000 or "DC_SCRIPT_ID = 'TPxi_Duplicates'" not in code or 'def act_find' not in code:
        return {'success': False, 'message': 'The download did not look like TPxi Duplicates, so nothing was changed.'}
    m = re.search(r"VERSION = '([\d.]+)'", code)
    model.WriteContentPython(name, code)
    return {'success': True, 'version': m.group(1) if m else '', 'message': 'Updated %s. Reloading.' % name}


def handle_post():
    if not can_use():
        print safe_json({'success': False, 'message': 'You need the Admin or Manager2 role.'})
        return
    if _param('action') == 'apply_update':
        print safe_json(apply_update() if _in_role('Admin') else {'success': False, 'message': 'Only an Admin can update this script.'})
        return
    try:
        fn = {'find': act_find, 'lookup': act_lookup, 'pair': act_pair, 'before': act_before, 'after': act_after,
              'verify': act_verify, 'notsame': act_notsame, 'detail': act_detail, 'values': act_values,
              'log': act_log, 'merges': act_merges}.get(_param('action'))
        res = fn() if fn else {'success': False, 'message': 'Unknown action.'}
        print safe_json(res)
    except Exception as e:
        print safe_json({'success': False, 'message': unicode(e)[:300]})


# ---------------------------------------------------------------- page
REDIRECT_JS = """<script>
if (location.pathname.indexOf('/PyScript/') > -1)
  location.replace(location.pathname.replace('/PyScript/', '/PyScriptForm/') + location.search);
</script>"""

CSS = """<style>
#mg [hidden]{display:none!important}
#mg .mgupdate{margin:0 0 12px;padding:10px 14px;border-radius:10px;background:#E4EEF7;color:#1F3A5F;font-size:14px}
#mg .mgupdate button{margin-left:6px}
#mg .busy{display:flex;align-items:center;gap:12px;font-size:16px;font-weight:600;color:#1F5A8A;padding:18px 4px}
#mg .spin{width:22px;height:22px;border-radius:50%;border:3px solid #C9D7E5;border-top-color:#1F5A8A;animation:mgspin .8s linear infinite;flex-shrink:0}
@keyframes mgspin{to{transform:rotate(360deg)}}
#mg{font-family:Inter,system-ui,-apple-system,sans-serif;color:#1B2220;max-width:1200px;margin:0 auto;padding:8px 4px 40px;font-size:15px}
#mg h1{font-size:28px;margin:0;font-weight:700}
#mg h2{font-size:19px;margin:0;font-weight:700}
#mg .eyebrow{font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#5B6461}
#mg .muted{color:#5B6461}
#mg .card{background:#fff;border:1px solid #E3E5DF;border-radius:16px;padding:16px 18px;display:flex;flex-direction:column;gap:10px;min-width:0}
#mg .row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
#mg .stack{display:flex;flex-direction:column;gap:14px}
#mg .btn{font:inherit;font-weight:600;font-size:14px;border-radius:10px;padding:9px 14px;cursor:pointer;border:2px solid #1F5A8A;background:#1F5A8A;color:#fff;text-decoration:none;display:inline-block}
#mg .btn.ghost{background:#fff;color:#1F5A8A;border-color:#C9D7E5}
#mg .btn.warn{background:#9B2C2C;border-color:#9B2C2C}
#mg .btn:disabled{opacity:.5;cursor:default}
#mg .chip{font:inherit;font-size:13px;font-weight:600;border-radius:999px;padding:6px 12px;border:2px solid #C9D7E5;background:#fff;color:#1F5A8A;cursor:pointer}
#mg .chip.on{background:#1F5A8A;border-color:#1F5A8A;color:#fff}
#mg .tag{font-size:12px;font-weight:600;border-radius:999px;padding:3px 9px;background:#E4EEF7;color:#1F5A8A;white-space:nowrap}
#mg .tag.warn{background:#FCEBD3;color:#7A4A00}
#mg .tag.bad{background:#FBE3E1;color:#9B2C2C}
#mg .pairs{display:flex;flex-direction:column}
#mg .pair{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr) minmax(0,1.1fr) auto;gap:14px;align-items:center;padding:12px 4px;border-bottom:1px solid #ECEDE8}
#mg .pair:last-child{border-bottom:0}
#mg .who b{display:block;font-size:15px}
#mg .who span{font-size:13px;color:#5B6461}
#mg .grid2{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:14px}
#mg .holds{display:grid;grid-template-columns:170px minmax(0,1fr) minmax(0,1fr);gap:6px 14px}
#mg .holds .h{font-weight:600}
#mg .colhead{font-weight:700;border-bottom:3px solid #D6D9D3;padding-bottom:6px}
#mg .colhead.keep{border-color:#1F5A8A}
#mg .fam{display:flex;flex-direction:column;gap:4px}
#mg .fam div{padding:4px 8px;border-radius:8px}
#mg .fam .self{font-weight:700}
#mg .fam .match{background:#E4EEF7}
#mg table.fields{border-collapse:collapse;width:100%}
#mg table.fields td,#mg table.fields th{padding:8px;border-bottom:1px solid #ECEDE8;text-align:left;vertical-align:top}
#mg table.fields th{font-size:12px;letter-spacing:1px;text-transform:uppercase;color:#5B6461}
#mg table.fields label{display:flex;flex-wrap:wrap;gap:4px 8px;align-items:center;cursor:pointer}
#mg table.fields td:first-child{width:24%;min-width:150px}
#mg table.fields label > span:first-of-type{flex:1 1 160px;min-width:0}
#mg table.fields tr.same{display:none}
#mg table.fields.all tr.same{display:table-row;color:#5B6461}
#mg .box{border:2px solid #E9C77F;background:#FFFBF2;border-radius:12px;padding:12px 14px}
#mg ul{margin:0;padding-left:20px;display:flex;flex-direction:column;gap:6px}
#mg .ok{color:#1E6B45;font-weight:600}
#mg .bad{color:#9B2C2C;font-weight:600}
#mg input[type=text]{font:inherit;padding:8px 10px;border:2px solid #D6D9D3;border-radius:10px}
#mg .scroll{overflow-x:auto}
#mg .linkbtn{font:inherit;font-weight:600;background:none;border:0;padding:0;color:#1F5A8A;cursor:pointer;text-align:left;text-decoration:underline;text-underline-offset:3px}
#mg .detail{grid-column:1/-1;background:#F5F7FA;border:1px solid #E1E8F0;border-radius:12px;padding:12px 14px;margin:2px 0 8px;display:flex;flex-direction:column;gap:10px}
#mg .detail .cols{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:16px}
#mg .detail .item{font-size:14px;padding:3px 0}
#mg .detail .shared{font-weight:700;color:#1F5A8A}
#mg .detail img{max-width:220px;max-height:220px;border-radius:10px;border:1px solid #E1E8F0}
#mg .detail .hl{background:#FDF1EF;border:1px solid #F0C9C3;border-radius:10px;padding:8px 10px;font-size:14px}
#mg .detail .ok2{background:#EAF4EE;border:1px solid #C7E2D1;border-radius:10px;padding:8px 10px;font-size:14px}
@media (max-width:800px){#mg .pair,#mg .grid2{grid-template-columns:minmax(0,1fr)}#mg .holds{grid-template-columns:110px minmax(0,1fr) minmax(0,1fr)}}
</style>"""

HTML = """
<div id="mg">
 <div class="mgupdate" id="mgupdate" hidden></div>
 <div id="mglist" class="stack">
  <div class="row"><div style="flex:1;min-width:240px"><div class="eyebrow">TPxi Duplicates</div><h1>Likely duplicates</h1></div>
   <span class="muted" id="mgcount"></span></div>
  <div class="row" id="mgtiers">
   <button type="button" class="chip on" data-tier="">All</button>
   <button type="button" class="chip" data-tier="strong">Strong</button>
   <button type="button" class="chip" data-tier="name">Name changed</button>
   <button type="button" class="chip" data-tier="check">No birthday</button>
   <button type="button" class="chip" data-tier="twins">Possibly twins</button>
   <button type="button" class="chip" data-tier="over">Record overwritten</button>
   <button type="button" class="chip" data-tier="parent">Possibly parent and child</button>
   <span class="muted" style="margin-left:8px">Newer record added</span>
   <select id="mgdays" class="chip"><option value="0">any time</option><option value="7">in the last 7 days</option><option value="30">in the last 30 days</option><option value="90">in the last 90 days</option><option value="365">in the last year</option></select>
  </div>
  <div class="muted">Strong: same birthday, and names alike (including nicknames and misspellings). Name changed: same birthday and first name,
   a different last name, and a shared email, cell or address. No birthday: names alike and a shared email or cell. Twins, Jr. and Sr., different genders,
   and pairs marked Not the same are left out. Nothing merges until you review a pair and press Merge.</div>
  <details class="card" id="mglogwrap"><summary style="font-weight:700;cursor:pointer">Merge log</summary><div id="mglog" class="stack" style="gap:8px;margin-top:8px"></div></details>
  <details class="card" id="mgsrcwrap"><summary style="font-weight:700;cursor:pointer">Where duplicates come from, and who keeps getting duplicated</summary><div id="mgsrc" class="stack" hidden></div></details>
  <div class="card"><div class="pairs" id="mgpairs"><div class="busy"><span class="spin"></span>Looking for likely duplicates across everyone in TouchPoint...</div></div></div>
  <div class="card">
   <h2>Compare any two people</h2>
   <div class="row"><input type="text" id="mgq" placeholder="Name or people id" style="min-width:240px"><button type="button" class="btn ghost" id="mgqgo">Find</button>
    <span class="muted" id="mgpicked"></span></div>
   <div id="mgqres" class="stack" style="gap:6px"></div>
  </div>
 </div>

 <div id="mgreview" class="stack" hidden>
  <div class="row"><button type="button" class="btn ghost" id="mgback1">&lsaquo; Likely duplicates</button></div>
  <div class="row"><h1 style="flex:1;min-width:240px" id="mgrtitle"></h1><button type="button" class="btn ghost" id="mgswap">Swap which one stays</button></div>
  <div id="mgrbody" class="stack"></div>
 </div>

 <div id="mgpreview" class="stack" hidden>
  <div class="row"><button type="button" class="btn ghost" id="mgback2">&lsaquo; Side by side</button></div>
  <div><h1>Before you merge</h1><div class="muted" id="mgpsub"></div></div>
  <div id="mgpbody" class="stack"></div>
 </div>

 <div id="mgdone" class="stack" hidden>
  <h1 id="mgdtitle">Merging...</h1>
  <div class="card" id="mgdlog"></div>
  <div class="row" id="mgdbtns" hidden><a class="btn ghost" id="mgdopen" target="_blank" rel="noopener">Open the person kept in TouchPoint</a>
   <button type="button" class="btn" id="mgdnext">Back to the list</button></div>
 </div>
 <div class="muted" style="margin-top:20px;font-size:12px">TPxi Duplicates %s &middot; Free from TPxi. <a href="https://github.com/bswaby/Touchpoint" target="_blank" rel="noopener">More free tools</a> &middot; <a href="https://tpxigo.com" target="_blank" rel="noopener">TPxi Go</a> brings TouchPoint to your phone and Outlook.</div>
</div>
"""

JS = r"""<script>
(function(){
  var base = location.pathname.replace('/PyScript/', '/PyScriptForm/');
  var PAIRS = [], TIER = '', CUR = null, FIELDS = [], PLAN = null, PICK = null;
  function $(id){ return document.getElementById(id); }
  function el(tag, attrs, kids){ var e = document.createElement(tag); attrs = attrs || {};
    for (var k in attrs){ if (k == 'text') e.textContent = attrs[k]; else if (k == 'html') e.innerHTML = attrs[k]; else if (k.slice(0,2) == 'on') e[k] = attrs[k]; else e.setAttribute(k, attrs[k]); }
    (kids || []).forEach(function(c){ if (c != null) e.appendChild(typeof c == 'string' ? document.createTextNode(c) : c); }); return e; }
  function show(id){ ['mglist','mgreview','mgpreview','mgdone'].forEach(function(s){ $(s).hidden = s != id; }); window.scrollTo(0, 0); }
  function busy(t){ return '<div class="busy"><span class="spin"></span><span>' + String(t).replace(/</g, '&lt;') + '</span></div>'; }
  function names(list){ var n = list.map(function(x){ return x.name + (x.age != null ? ' (' + x.age + ')' : ''); });
    return n.length < 2 ? n.join('') : n.slice(0, -1).join(', ') + ' and ' + n[n.length - 1]; }
  function plural(n, one, many){ return n + ' ' + (n == 1 ? one : many); }
  function post(d, cb){
    var b = []; for (var k in d) b.push(encodeURIComponent('mg_' + k) + '=' + encodeURIComponent(d[k]));
    var x = new XMLHttpRequest(); x.open('POST', base, true);
    x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    x.onload = function(){ var t = x.responseText, i = t.indexOf('{'), j = t.lastIndexOf('}'), r;
      try { r = JSON.parse(t.slice(i, j + 1)); } catch(e){ r = {success: false, message: 'Unexpected reply from TouchPoint (signed out?).'}; }
      cb(r); };
    x.onerror = function(){ cb({success: false, message: 'Could not reach TouchPoint.'}); };
    x.send(b.join('&'));
  }

  // ---- the list
  function load(){
    $('mgpairs').innerHTML = busy('Looking for likely duplicates across everyone in TouchPoint (a few seconds)...'); $('mgcount').textContent = 'Searching...';
    post({action: 'find', days: $('mgdays').value}, function(r){
      if (!r.success) { $('mgpairs').innerHTML = '<span class="bad"></span>'; $('mgpairs').firstChild.textContent = r.message; return; }
      PAIRS = r.pairs; var c = r.counts; drawSources(r);
      $('mgcount').textContent = plural(PAIRS.filter(function(p){ return !p.twins && !p.over && !p.parent; }).length, 'pair', 'pairs') + (c.twins ? ', ' + c.twins + ' possibly twins' : '') + (c.over ? ', ' + c.over + ' with a record overwritten' : '') + (c.parent ? ', ' + c.parent + ' possibly parent and child' : '') + (r.capped ? ' (the first 400)' : '');
      Array.prototype.forEach.call(document.querySelectorAll('#mgtiers .chip[data-tier]'), function(b){
        var t = b.getAttribute('data-tier'), n = t ? c[t] || 0 : PAIRS.length;
        if (!t) n = PAIRS.filter(function(p){ return !p.twins && !p.over && !p.parent; }).length;
        else if (t == 'strong' || t == 'name' || t == 'check') n = PAIRS.filter(function(p){ return !p.twins && !p.over && !p.parent && p.tier == t; }).length;
        b.textContent = ({'': 'All', strong: 'Strong', name: 'Name changed', check: 'No birthday', twins: 'Possibly twins', over: 'Record overwritten', parent: 'Possibly parent and child'})[t] + ' (' + n + ')'; });
      draw();
    });
  }
  function who(p, note){ return el('div', {'class': 'who'}, [el('b', {text: p.name + (p.age != null ? ', ' + p.age : '')}),
    el('span', {text: '#' + p.pid + ' \u00b7 added ' + p.created + (note ? ' \u00b7 ' + note : '')})]); }
  function hist(p){ var h = []; if (p.family) h.push('family of ' + (p.family + 1)); if (p.gifts) h.push(plural(p.gifts, 'gift', 'gifts')); if (p.att) h.push(p.att + ' attended');
    if (p.login) h.push('login'); if (p.how) h.push(p.how); return h.join(' \u00b7 ') || 'little history'; }
  function draw(){
    var box = $('mgpairs'); box.innerHTML = '';
    var list = PAIRS.filter(function(p){ return TIER == 'over' ? p.over : (TIER == 'twins' ? p.twins && !p.over : (TIER == 'parent' ? p.parent && !p.over && !p.twins : (!p.twins && !p.over && !p.parent && (!TIER || p.tier == TIER)))); });
    if (TIER == 'parent') box.appendChild(el('div', {'class': 'muted', style: 'padding:4px 4px 10px', text: 'One record has no birthday and was added as an adult while the other was a child, and they share no involvement: usually a parent and child with the same name and the same email or phone. Mark them Not the same.'}));
    if (TIER == 'over') box.appendChild(el('div', {'class': 'muted', style: 'padding:4px 4px 10px', text: 'One record in each pair had its first name changed together with its gender or birthday: usually a parent who signed in to register a child, and the form saved the child onto the parent. These are two people. Change the overwritten record back in TouchPoint, then mark the pair Not the same.'}));
    if (TIER == 'twins') box.appendChild(el('div', {'class': 'muted', style: 'padding:4px 4px 10px', text: 'Same birthday and different first names, and either entered together or present at the same meetings. Usually twins: mark them Not the same so they stop coming back.'}));
    if (!list.length) { box.appendChild(el('span', {'class': 'muted', text: 'No likely duplicates here.'})); return; }
    box.appendChild(el('div', {'class': 'pair', style: 'padding-top:0'}, [el('span', {'class': 'eyebrow', text: 'Suggested to stay'}),
      el('span', {'class': 'eyebrow', text: 'Probably the same person'}), el('span', {'class': 'eyebrow', text: 'Why'}), el('span')]));
    list.forEach(function(p){
      var why = p.why.slice(); if (p.keep.merged || p.lose.merged) why.push('Merged before');
      if (p.twins) why.push('Possibly twins: ' + p.twinWhy);
      if (p.parent) why.push('Possibly parent and child: ' + p.parentWhy);
      if (p.over) why.push('Overwritten: #' + p.over.pid + ' was ' + p.over.from + ' until ' + p.over.when);
      // both live with other people in different families (a grown child with their own household, say): look closely
      if (p.keep.family && p.lose.family && why.indexOf('Same family') < 0) why.push('Both in families');
      var tags = el('div', {'class': 'row', style: 'gap:6px'}, why.map(function(w){ return el('span', {'class': 'tag' + (/changed|No birthday|Both in families|Possibly twins|not twins|January 1/.test(w) ? ' warn' : (/Merged before|Overwritten|parent and child/.test(w) ? ' bad' : '')), text: w}); }));
      var ns = el('button', {type: 'button', 'class': 'btn ghost', text: 'Not the same'});
      ns.onclick = function(){ if (!ns.dataset.sure) { ns.dataset.sure = '1'; ns.textContent = 'Tap again: not the same'; return; }
        ns.disabled = true; post({action: 'notsame', a: p.keep.pid, b: p.lose.pid}, function(r){
          if (!r.success) { ns.textContent = r.message; return; }
          PAIRS = PAIRS.filter(function(x){ return x !== p; }); draw(); }); };
      box.appendChild(el('div', {'class': 'pair'}, [who(p.keep, hist(p.keep)), who(p.lose, hist(p.lose)), tags,
        el('div', {'class': 'row', style: 'justify-content:flex-end'}, [el('button', {type: 'button', 'class': 'btn', text: 'Review', onclick: function(){ review(p.keep.pid, p.lose.pid); }}), ns])]));
    });
  }
  function drawSources(r){
    var box = $('mgsrc'); box.innerHTML = ''; box.hidden = false;
    box.appendChild(el('h2', {text: 'Where duplicates come from'}));
    box.appendChild(el('div', {'class': 'muted', text: 'How the newer record in each pair above was added. Fixing the biggest source beats merging forever.'}));
    var g = el('div', {'class': 'grid2'});
    g.appendChild(el('div', {'class': 'stack', style: 'gap:4px'}, [el('b', {text: 'Added through'})].concat((r.how || []).slice(0, 10).map(function(x){ return el('div', {text: x[1] + '  ' + x[0]}); }))));
    g.appendChild(el('div', {'class': 'stack', style: 'gap:4px'}, [el('b', {text: 'Added by'}),
      el('div', {'class': 'muted', style: 'font-size:13px', text: 'Whoever adds the most people will have the most duplicates; the rate is the fair number. Most come from TouchPoint\'s add search missing a spelling (Rachel and Rachael).'})].concat((r.by || []).length ? r.by.map(function(x){
        return el('div', {text: x[0] + ': ' + plural(x[1], 'duplicate', 'duplicates') + (x[2] ? ' of ' + x[2].toLocaleString() + ' people added (' + (x[1] * 100 / x[2]).toFixed(1) + '%)' : '')}); }) : [el('div', {'class': 'muted', text: 'Not recorded (forms, check-in and the app add people without a name here)'})])));
    box.appendChild(g);
    var rep = r.repeats || [];
    box.appendChild(el('h2', {style: 'margin-top:8px', text: 'Keeps getting duplicated'}));
    if (!rep.length) box.appendChild(el('div', {'class': 'muted', text: 'No one has been merged more than once.'}));
    else {
      box.appendChild(el('div', {'class': 'muted', text: plural(rep.length, 'person has', 'people have') + ' been merged two or more times. Something keeps adding them again (a check-in, a form, the app).'}));
      box.appendChild(el('div', {'class': 'stack', style: 'gap:4px'}, rep.map(function(x){ return el('div', {'class': 'row'}, [el('a', {href: '/Person2/' + x.pid, target: '_blank', rel: 'noopener', text: x.name + ' #' + x.pid}),
        el('span', {'class': 'muted', text: 'merged ' + x.n + ' times, ' + x.first + ' to ' + x.last})]); })));
    }
  }
  Array.prototype.forEach.call(document.querySelectorAll('#mgtiers .chip[data-tier]'), function(b){
    b.onclick = function(){ TIER = b.getAttribute('data-tier');
      Array.prototype.forEach.call(document.querySelectorAll('#mgtiers .chip[data-tier]'), function(x){ x.classList.toggle('on', x === b); }); draw(); }; });
  $('mgdays').onchange = load;

  // ---- compare any two by hand
  $('mgqgo').onclick = function(){ var t = $('mgq').value.trim(); if (!t) return;
    post({action: 'lookup', q: t}, function(r){ var box = $('mgqres'); box.innerHTML = '';
      if (!r.success) { box.textContent = r.message; return; }
      if (!r.people.length) box.appendChild(el('span', {'class': 'muted', text: 'No one found.'}));
      r.people.forEach(function(p){ box.appendChild(el('div', {'class': 'row'}, [who(p),
        el('button', {type: 'button', 'class': 'btn ghost', text: PICK ? 'Compare with ' + PICK.name : 'Pick', onclick: function(){
          if (!PICK) { PICK = p; $('mgpicked').textContent = 'Picked ' + p.name + ' (#' + p.pid + '). Now find the other one.'; box.innerHTML = ''; $('mgq').value = ''; return; }
          if (PICK.pid == p.pid) return; var a = PICK; PICK = null; $('mgpicked').textContent = ''; review(a.pid, p.pid); }})])); }); }); };
  $('mgq').onkeydown = function(e){ if (e.key == 'Enter') { e.preventDefault(); $('mgqgo').click(); } };

  // ---- side by side
  function review(keep, lose){
    CUR = {keep: keep, lose: lose}; FIELDS = []; show('mgreview');
    $('mgrtitle').textContent = 'Comparing two records'; $('mgrbody').innerHTML = busy('Loading both records, their families and history...');
    post({action: 'pair', keep: keep, lose: lose}, function(r){
      if (!r.success) { $('mgrbody').innerHTML = ''; $('mgrbody').appendChild(el('span', {'class': 'bad', text: r.message})); return; }
      CUR.data = r; drawReview(r); loadFields();
    });
  }
  $('mgswap').onclick = function(){ if (CUR) review(CUR.lose, CUR.keep); };
  $('mgback1').onclick = function(){ show('mglist'); };
  function line(label, a, b){ return [el('span', {'class': 'h', text: label}), el('span', {text: a == null || a === '' ? 'none' : String(a)}), el('span', {text: b == null || b === '' ? 'none' : String(b)})]; }
  function drawReview(r){
    var k = r.keep, l = r.lose, kc = k.counts, lc = l.counts;
    $('mgrtitle').textContent = k.name + ' #' + k.pid + ' and ' + l.name + ' #' + l.pid;
    var body = $('mgrbody'); body.innerHTML = '';
    var overs = [];
    [k, l].forEach(function(p){ (p.identity || []).forEach(function(c){ overs.push({p: p, c: c}); }); });
    if (overs.length) {
      var ob = el('div', {'class': 'box', style: 'border-color:#C2453B;background:#FDF1EF'}, [el('h2', {text: 'Probably two people: a record was overwritten'})]);
      overs.forEach(function(o){ var c = o.c;
        // what happened, in a sentence
        var what = 'What happened: ' + (c.self
          ? 'on ' + c.when + ', ' + c.from + (c.login ? ' signed in at ' + c.login + ' and' : '') + (c.reg ? ' finished the ' + c.reg + ' registration' : ' saved a form') + ' at ' + c.at +
            ' with ' + c.to + '\'s name' + (c.genderTo ? ', gender' : '') + (c.birthTo ? ' and birthday' : '') + ' on ' + (c.genderFrom == 'Female' ? 'her' : (c.genderFrom == 'Male' ? 'his' : 'their')) + ' own record. ' +
            'Most likely ' + c.from + ' was chosen on the form instead of adding ' + c.to + ', and ' + c.to + '\'s details were typed over ' + c.from + '\'s.'
          : 'on ' + c.when + ' at ' + c.at + (c.by ? ', ' + c.by : ', someone') + (c.reg ? ' (during the ' + c.reg + ' registration)' : '') + ' changed this record from ' + c.from + ' to ' + c.to +
            ', with a different gender or birthday: a different person written over it.');
        ob.appendChild(el('div', {style: 'margin-top:6px', text: '#' + o.p.pid + ' was ' + c.from + (c.genderFrom ? ' (' + c.genderFrom + (c.birthFrom ? ', born ' + c.birthFrom : '') + ')' : '') +
          ' until ' + c.when + '. It now reads ' + c.to + (c.genderTo ? ' (' + c.genderTo + ')' : '') + (c.birthTo ? ', born ' + c.birthTo : '') + '.'}));
        ob.appendChild(el('div', {style: 'margin-top:4px', text: what}));
        var tl = el('div', {'class': 'detail', style: 'margin-top:8px'}); tl.hidden = true;
        var tb = el('button', {type: 'button', 'class': 'linkbtn', style: 'margin-top:6px', text: 'Show the timeline for #' + o.p.pid});
        tb.onclick = function(){ tl.hidden = !tl.hidden; if (!tl.hidden && !tl.dataset.loaded) { tl.innerHTML = busy('Loading the timeline...');
          post({action: 'detail', kind: 'timeline', keep: k.pid, lose: l.pid}, function(r2){ tl.innerHTML = ''; tl.dataset.loaded = '1';
            if (!r2.success) { tl.appendChild(el('span', {'class': 'bad', text: r2.message})); return; }
            drawTimeline(tl, o.p.pid == k.pid ? r2.keep : r2.lose, c.when); }); } };
        ob.appendChild(tb); ob.appendChild(tl); });
      ob.appendChild(el('div', {style: 'margin-top:8px;font-weight:600', text: 'These are two people. Merging would delete the other record and leave this one as the wrong person. Change the overwritten record back in TouchPoint (the Changes tab has the old values), then mark this pair Not the same.'}));
      var nsb = el('button', {type: 'button', 'class': 'btn', text: 'Mark Not the same'});
      nsb.onclick = function(){ nsb.disabled = true; post({action: 'notsame', a: k.pid, b: l.pid}, function(r2){ if (!r2.success) { nsb.textContent = r2.message; return; }
        PAIRS = PAIRS.filter(function(x){ return !(x.keep.pid == k.pid && x.lose.pid == l.pid) && !(x.keep.pid == l.pid && x.lose.pid == k.pid); }); show('mglist'); draw(); }); };
      ob.appendChild(el('div', {'class': 'row', style: 'margin-top:8px'}, overs.map(function(o){ return el('a', {'class': 'btn ghost', href: '/Person2/' + o.p.pid, target: '_blank', rel: 'noopener', text: 'Open #' + o.p.pid + ' to change it back'}); }).concat([nsb])));
      body.appendChild(ob);
      CUR.overwritten = true;
    } else CUR.overwritten = false;
    // more than two records for one person (Dan Sharp had three): merge one pair at a time
    CUR.also = PAIRS.filter(function(p){ var a = p.keep.pid, b = p.lose.pid;
      return !((a == k.pid && b == l.pid) || (a == l.pid && b == k.pid)) && (a == k.pid || b == k.pid || a == l.pid || b == l.pid); });
    if (CUR.also.length) body.appendChild(el('div', {'class': 'box'}, [el('b', {text: 'This person may have more than two records. '}),
      'Also on the list: ' + CUR.also.map(function(p){ return p.keep.name + ' #' + p.keep.pid + ' and ' + p.lose.name + ' #' + p.lose.pid; }).join('; ') +
      '. Merge this pair first; afterward this page offers the next comparison with the record kept.']));
    var holds = el('div', {'class': 'holds'}, [el('span', {'class': 'eyebrow', style: 'align-self:end', text: 'What each holds'}),
      el('div', {'class': 'colhead keep'}, [el('div', {'class': 'eyebrow', style: 'color:#1F5A8A', text: 'Stays'}), el('b', {text: k.name + ' #' + k.pid})]),
      el('div', {'class': 'colhead'}, [el('div', {'class': 'eyebrow', text: 'Moved into them, then deleted'}), el('b', {text: l.name + ' #' + l.pid})])]);
    function add(cells, kind){
      if (kind) {   // a row that opens to show what is in it, for both records
        var label = cells[0].textContent, det = el('div', {'class': 'detail'}); det.hidden = true;
        var btn = el('button', {type: 'button', 'class': 'linkbtn', text: label, title: 'Show what is in it'});
        btn.onclick = function(){ det.hidden = !det.hidden; if (!det.hidden && !det.dataset.loaded) loadDetail(kind, det); };
        cells[0] = el('span', {'class': 'h'}, [btn]);
        cells.forEach(function(c){ holds.appendChild(c); }); holds.appendChild(det); return;
      }
      cells.forEach(function(c){ holds.appendChild(c); });
    }
    add(line('Added', k.created + (k.createdBy ? ' by ' + k.createdBy : ''), l.created + (l.createdBy ? ' by ' + l.createdBy : '')));
    add(line('How', k.how || 'not recorded', l.how || 'not recorded'));
    function mergedTxt(p){ return p.merged.length ? plural(p.merged.length, 'time', 'times') + ': ' + p.merged.map(function(m){ return m.when + ' from #' + m.from + (m.who ? ' by ' + m.who : ''); }).join('; ') : 'No'; }
    add(line('Merged before', mergedTxt(k), mergedTxt(l)));
    add(line('Birthday', k.birthday + (k.age != null ? ' (' + k.age + ')' : ''), l.birthday + (l.age != null ? ' (' + l.age + ')' : '')));
    add(line('Member status', k.member, l.member));
    add(line('Membership history', k.membership.join('; ') || 'none', l.membership.join('; ') || 'none'));
    if (kc.gifts == null) add(line('Gifts', 'hidden (needs the Finance role)', 'hidden; the merge moves every gift'));
    else add(line('Gifts', kc.gifts + (kc.lastGift ? ', last ' + kc.lastGift : ''), lc.gifts + (lc.lastGift ? ', last ' + lc.lastGift : '')));
    add(line('Attended', kc.att + (kc.lastAtt ? ', last ' + kc.lastAtt : ''), lc.att + (lc.lastAtt ? ', last ' + lc.lastAtt : '')), 'att');
    add(line('Involvements', kc.orgs + (k.orgs.length ? ': ' + k.orgs.map(function(o){ return o.name; }).join(', ') : ''), lc.orgs + (l.orgs.length ? ': ' + l.orgs.map(function(o){ return o.name; }).join(', ') : '')), 'orgs');
    add(line('Login', kc.logins ? 'Yes' + (kc.lastLogin ? ', last used ' + kc.lastLogin : '') : 'None', lc.logins ? 'Yes' + (lc.lastLogin ? ', last used ' + lc.lastLogin : '') : 'None'), 'login');
    add(line('Registrations', kc.regs, lc.regs));
    add(line('Background checks', kc.checks, lc.checks));
    add(line('Notes and tasks', kc.notes, lc.notes), 'notes');
    add(line('Tags', kc.tags + (kc.flags ? ' + ' + plural(kc.flags, 'status flag', 'status flags') : ''), lc.tags + (lc.flags ? ' + ' + plural(lc.flags, 'status flag', 'status flags') : '')), 'tags');
    add(line('Photo', k.photo ? 'Yes' : 'None', l.photo ? 'Yes' : 'None'), 'photo');
    var V = r.values, MD = r.medical;
    add(line('Extra values', V.keep.length + (V.conflicts.length ? ', ' + plural(V.conflicts.length, 'differs', 'differ') + ' from the other' : ''), V.lose.length + (V.moves.length ? ', ' + V.moves.length + ' only here' : '')), 'evs');
    function medTxt(side){ var a = MD.filter(function(m){ return m.field == 'MedicalDescription'; })[0], e = MD.filter(function(m){ return m.field == 'Emcontact'; })[0];
      var parts = []; if (a && a[side]) parts.push('allergies: ' + a[side]); if (e && e[side]) parts.push('call ' + e[side]); return parts.join('; ') || 'none'; }
    add(line('Medical and emergency', medTxt('keep'), medTxt('other')), 'med');
    add(line('Timeline', 'added ' + k.created + (k.identity.length ? ', changed to another person ' + k.identity[k.identity.length - 1].when : ''), 'added ' + l.created + (l.identity.length ? ', changed to another person ' + l.identity[l.identity.length - 1].when : '')), 'timeline');
    var loss = r.loss, hint = null;
    // the person kept can move into the family the other record is in (TouchPoint's Person > Family > Add)
    CUR.joinFamily = !!(loss.leavesFamily.length && loss.keepAlone && !loss.olderFamily);
    function joinBox(on){ var c = el('input', {type: 'checkbox', style: 'width:20px;height:20px'}); c.checked = on; c.onchange = function(){ CUR.joinFamily = c.checked; if ($('mgfields') && FIELDS.length) loadFields(); };
      return el('label', {'class': 'row', style: 'font-weight:600;margin-top:8px'}, [c, 'Move #' + k.pid + ' into the family with ' + names(loss.leavesFamily) + ' right after the merge']); }
    if (loss.leavesFamily.length && loss.keepAlone) hint = el('div', {'class': 'box', style: 'border-color:#9CC3E4;background:#F2F8FD'}, [
      el('b', {text: '#' + k.pid + ' is the original record (added ' + k.created + '). '}),
      '#' + l.pid + ' is the one in the family with ' + names(loss.leavesFamily) + '. Keeping the original keeps its record number, added date and membership history, and ' + k.name + ' moves into that family.',
      loss.olderFamily ? el('div', {'class': 'warn', style: 'margin-top:6px', text: 'Everyone in that family is 45 or more years older than ' + k.name + ' (probably grandparents), so the move is not checked. Check it only if ' + k.name + ' lives with them.'}) : null,
      joinBox(!loss.olderFamily)]);
    else if (loss.leavesFamily.length) {
      // both records live with other people: which household is this person in now?
      var headNote = loss.headless ? ' ' + names(loss.leavesFamily) + ' would be left in a family with no head of household unless brought along (the next screen offers it).' : '';
      var keepOthersNote = loss.keepOthers.length ? ' with ' + names(loss.keepOthers) : '';
      var child = k.age != null && k.age < 18;
      hint = el('div', {'class': 'box'}, [el('b', {text: child ? 'A child in two homes? ' : 'Both records live with other people. Which household is ' + k.name + ' in now? '}),
        child ? el('div', {style: 'margin-top:6px', text: 'Each home entered ' + k.name + ': separated parents, or grandparents who bring them. TouchPoint keeps one family, so choose the home ' + k.name + ' lives in most (ask someone who knows the family), and link the other as a related family in TouchPoint afterward.'}) : null,
        el('div', {style: 'margin-top:6px', text: 'Keeping #' + k.pid + ' (as now): ' + k.name + ' stays in the family' + keepOthersNote + ', and the family with ' + names(loss.leavesFamily) + ' no longer includes them.' + headNote}),
        el('div', {style: 'margin-top:4px', text: 'Keeping #' + l.pid + ' instead: ' + k.name + ' stays with ' + names(loss.leavesFamily) + (loss.keepOthers.length ? ', and ' + names(loss.keepOthers) + ' can be brought into that family in the same step' : '') + '.'}),
        joinBox(false),
        loss.keepOthers.length ? el('div', {'class': 'muted', style: 'margin-top:4px', text: 'Moving the original keeps its record number and membership history; ' + names(loss.keepOthers) + ' can come along (the next screen asks).'}) : null,
        el('div', {'class': 'row', style: 'margin-top:8px'}, [el('button', {type: 'button', 'class': 'btn ghost', text: 'Keep #' + l.pid + ' instead', onclick: function(){ review(l.pid, k.pid); }}),
          el('span', {'class': 'muted', text: 'A kiosk or form that missed the family often makes the newer one; a person who moved out also looks like this.'})])]);
    }
    else if (lc.gifts || lc.logins || lc.att) hint = el('div', {'class': 'muted', text: '#' + l.pid + "'s " + [lc.gifts ? 'giving' : '', lc.att ? 'attendance' : '', lc.logins ? 'login' : ''].filter(Boolean).join(', ') + ' and the rest of its history move to #' + k.pid + ' in the merge, so nothing is lost by keeping #' + k.pid + '.'});
    var repl = r.medical.filter(function(m){ return m.what == 'replaced'; }), medWarn = null;
    if (repl.length) medWarn = el('div', {'class': 'box', style: 'border-color:#C2453B;background:#FDF1EF'}, [el('b', {text: 'Medical and emergency info will change. '}),
      'When both records have it, TouchPoint replaces the kept record\'s with the deleted one\'s: ' + repl.map(function(m){ return m.label + ' on #' + k.pid + ' becomes "' + m.other + '" (was "' + m.keep + '")'; }).join('; ') +
      '. If #' + k.pid + '\'s is the current one, change it back on its page after merging (the result links it).']);
    body.appendChild(el('div', {'class': 'card scroll'}, [holds, hint, medWarn]));

    // families
    var matched = {}, mkind = {}, byPid = {}; (r.keepFamily || []).concat(r.loseFamily || []).forEach(function(p){ byPid[p.pid] = p; });
    (r.familyMatches || []).forEach(function(m){ matched[m[0]] = m[1]; matched[m[1]] = m[0]; mkind[m[0]] = mkind[m[1]] = m[2] || 'same'; });
    function famList(list){ return el('div', {'class': 'fam'}, list.map(function(p){
      return el('div', {'class': (p.self ? 'self' : '') + (matched[p.pid] ? ' match' : '')}, [p.name + (p.age != null ? ', ' + p.age : '') + (p.position ? ' \u00b7 ' + p.position : '') + (matched[p.pid] ? ' \u00b7 ' + (mkind[p.pid] == 'maybe' ? 'maybe on both' : 'on both') : '')]); })); }
    function bdTxt(p){ return p.bm && p.bd ? p.bm + '/' + p.bd + (p.by ? '/' + p.by : '') : 'no birthday'; }
    // the others in both families are duplicates too: compare them next (their own pair, a click away)
    var cmpRows = (r.familyMatches || []).map(function(m){ var a = byPid[m[0]], b = byPid[m[1]]; if (!a || !b) return null;
      return el('div', {'class': 'row'}, [el('span', {style: 'flex:1;min-width:220px', text: a.name + ' #' + a.pid + ' and ' + b.name + ' #' + b.pid + (m[2] == 'maybe' ? ' (' + bdTxt(a) + ' and ' + bdTxt(b) + ')' : '')}),
        el('button', {type: 'button', 'class': 'btn ghost', text: 'Compare these two', onclick: function(){ review(a.pid, b.pid); }})]); }).filter(Boolean);
    function relList(list){ return list.length ? el('div', {'class': 'muted', text: 'Related: ' + list.map(function(x){ return x.name + (x.rel ? ' (' + x.rel + ')' : ''); }).join(', ')}) : null; }
    var famCard;
    if (r.sameFamily) famCard = el('div', {'class': 'card'}, [el('h2', {text: 'Their family'}), el('div', {'class': 'muted', text: 'Both records are in the same family.'}), famList(r.keepFamily), relList(r.keepRelated)]);
    else {
      var nm = (r.familyMatches || []).filter(function(m){ return m[2] != 'maybe'; }).length, nmb = (r.familyMatches || []).length - nm;
      famCard = el('div', {'class': 'card'}, [el('h2', {text: 'Their families'}),
        el('div', {'class': nm || nmb ? 'ok' : 'muted', text: (nm ? plural(nm, 'other person is', 'other people are') + ' in both families (same name and birthday): likely one family entered twice. ' : '') +
          (nmb ? plural(nmb, 'other person is', 'other people are') + ' probably in both too: the same first name, with a birthday a digit off or missing, so the list above cannot find them. ' : '') +
          (nm || nmb ? 'Merge this pair first, then compare them (the result screen offers it again).' : 'No one else matches between the two families.')}),
        cmpRows.length ? el('div', {'class': 'stack', style: 'gap:6px'}, cmpRows) : null,
        el('div', {'class': 'grid2'}, [el('div', {'class': 'stack', style: 'gap:6px'}, [el('b', {text: k.name + ' #' + k.pid + "'s family (stays)"}), famList(r.keepFamily), relList(r.keepRelated)]),
          el('div', {'class': 'stack', style: 'gap:6px'}, [el('b', {text: l.name + ' #' + l.pid + "'s family"}), famList(r.loseFamily), relList(r.loseRelated)])])]);
    }
    body.appendChild(famCard);
    CUR.evPick = {};
    if (V.conflicts.length) {
      var et = el('table', {'class': 'fields'}, [el('tr', {}, [el('th', {text: 'Extra value'}), el('th', {text: 'Keep #' + k.pid + "'s"}), el('th', {text: 'Use #' + l.pid + "'s"})])]);
      V.conflicts.forEach(function(c){
        var pick = c.keep ? 'keep' : 'other'; CUR.evPick[c.field] = c.settable ? pick : 'both';
        function opt(v, value){ var i = el('input', {type: 'radio', name: 'mgev_' + c.field, value: v}); if (pick == v) i.checked = true; i.onchange = function(){ CUR.evPick[c.field] = v; };
          return el('label', {}, [i, el('span', {text: value || '(blank)'})]); }
        et.appendChild(el('tr', {}, [el('td', {}, [el('b', {text: c.field})]),
          c.settable ? el('td', {}, [opt('keep', c.keep)]) : el('td', {text: c.keep}), c.settable ? el('td', {}, [opt('other', c.other)]) : el('td', {}, [el('span', {text: c.other}), el('div', {'class': 'muted', style: 'font-size:12px', text: 'Kept as ' + c.field + '_mv (this kind of value cannot be set here)'})])]));
      });
      body.appendChild(el('div', {'class': 'card'}, [el('h2', {text: 'Extra values on both records'}),
        el('div', {'class': 'muted', text: 'TouchPoint keeps #' + k.pid + "'s value and saves #" + l.pid + "'s as a second field named with _mv, which reports never read. Choose which value #" + k.pid + ' should have; the other is removed.' +
          (V.dups.length ? ' ' + plural(V.dups.length, 'other field has', 'other fields have') + ' the same value on both; their extra copies are removed too.' : '')}),
        el('div', {'class': 'scroll'}, [et])]));
    }
    body.appendChild(el('div', {'class': 'card', id: 'mgfields', html: '<h2>Where they differ</h2>' + busy("Reading TouchPoint's merge screen for this pair...")}));
    body.appendChild(el('div', {'class': 'row'}, [el('span', {'class': 'muted', style: 'flex:1;min-width:240px', text: 'Nothing changes yet. Next shows what TouchPoint will move, what it will drop, and what this page saves first.'}),
      el('a', {'class': 'btn ghost', href: '/Merge/' + l.pid + '/' + k.pid, target: '_blank', rel: 'noopener', text: "Open TouchPoint's merge instead"}),
      el('button', {type: 'button', 'class': 'btn', id: 'mgnext', disabled: 'disabled', text: 'Next: what happens', onclick: preview})]));
  }

  // ---- what is inside a row (photo, attendance, involvements, notes, tags, login)
  function loadDetail(kind, det){
    var k = CUR.data.keep, l = CUR.data.lose;
    if (kind == 'evs' || kind == 'med') { localDetail(kind, det); return; }
    det.innerHTML = busy('Loading...');
    post({action: 'detail', kind: kind, keep: k.pid, lose: l.pid}, function(r){
      det.innerHTML = '';
      if (!r.success) { det.appendChild(el('span', {'class': 'bad', text: r.message})); return; }
      det.dataset.loaded = '1';
      function col(title, kids){ return el('div', {'class': 'stack', style: 'gap:2px'}, [el('b', {text: title})].concat(kids && kids.length ? kids : [el('span', {'class': 'muted', text: 'None'})])); }
      function item(text, shared){ return el('div', {'class': 'item' + (shared ? ' shared' : ''), text: text + (shared ? ' (on both)' : '')}); }
      var a = '#' + k.pid + ' (stays)', b = '#' + l.pid;
      if (kind == 'photo') {
        det.appendChild(el('div', {'class': 'cols'}, [col(a, r.keep ? [el('img', {src: r.keep, alt: 'Photo of #' + k.pid})] : []), col(b, r.lose ? [el('img', {src: r.lose, alt: 'Photo of #' + l.pid})] : [])]));
      } else if (kind == 'att') {
        var sameFirst = (k.first || '').toLowerCase() == (l.first || '').toLowerCase();
        if (r.sameN) det.appendChild(el('div', {'class': sameFirst ? 'ok2' : 'hl', text: 'Both records were marked present at the same meeting ' + plural(r.sameN, 'time', 'times') + ' (' + r.same.map(function(x){ return x.d + ' ' + x.org; }).join('; ') + '). ' +
          (sameFirst ? 'With the same first name, that is usually one person on the roll twice (both records were checked), which a merge fixes.' : 'One person cannot be in two places, and the first names differ: these may be two different people, such as twins.')}));
        else det.appendChild(el('div', {'class': 'ok2', text: 'The two records were never at the same meeting together.'}));
        function side(x){ return [el('div', {'class': 'muted', text: 'Most attended'})].concat(x.top.map(function(o){ return item(o.org + ': ' + o.n + ', last ' + o.last); }))
          .concat([el('div', {'class': 'muted', style: 'margin-top:6px', text: 'Most recent'})]).concat(x.recent.map(function(o){ return item(o.d + ' ' + o.org); })); }
        det.appendChild(el('div', {'class': 'cols'}, [col(a, r.keep.top.length ? side(r.keep) : []), col(b, r.lose.top.length ? side(r.lose) : [])]));
      } else if (kind == 'orgs') {
        var ko = {}, lo = {}; r.keep.forEach(function(o){ ko[o.oid] = 1; }); r.lose.forEach(function(o){ lo[o.oid] = 1; });
        var both = r.keep.filter(function(o){ return lo[o.oid]; }).length;
        det.appendChild(el('div', {'class': 'muted', text: both ? plural(both, 'involvement is', 'involvements are') + ' on both records; TouchPoint keeps one membership.' : 'No involvement is on both records.'}));
        function orgs(list, other){ return list.map(function(o){ return item(o.org + (o.as ? ', ' + o.as : '') + (o.since ? ', since ' + o.since : ''), other[o.oid]); }); }
        det.appendChild(el('div', {'class': 'cols'}, [col(a, orgs(r.keep, lo)), col(b, orgs(r.lose, ko))]));
      } else if (kind == 'notes') {
        function notes(x){ var out = x.items.map(function(n){ return el('div', {'class': 'item'}, [el('b', {text: n.d + ' ' + (n.note ? 'Note' : 'Task') + (n.by ? ' by ' + n.by : '') + ': '}), n.text || '(no text)']); });
          if (x.hidden) out.push(el('div', {'class': 'muted', text: plural(x.hidden, 'more is', 'more are') + ' limited to a role you do not have.'})); return out; }
        det.appendChild(el('div', {'class': 'cols'}, [col(a, notes(r.keep)), col(b, notes(r.lose))]));
        det.appendChild(el('div', {'class': 'muted', text: 'Notes and tasks about either record move to #' + k.pid + ' in the merge.'}));
      } else if (kind == 'tags') {
        var kt = {}, lt = {}; r.keep.forEach(function(x){ kt[x.id] = 1; }); r.lose.forEach(function(x){ lt[x.id] = 1; });
        function tags(list, other){ return list.map(function(x){ return item((x.flag ? 'Status flag: ' : '') + x.name + (x.owner && !x.flag ? ' (' + x.owner + ')' : ''), other[x.id]); }); }
        det.appendChild(el('div', {'class': 'cols'}, [col(a, tags(r.keep, lt)), col(b, tags(r.lose, kt))]));
      } else if (kind == 'timeline') {
        var c2 = el('div', {'class': 'cols'}), ca = el('div', {'class': 'stack', style: 'gap:2px'}, [el('b', {text: a})]), cb = el('div', {'class': 'stack', style: 'gap:2px'}, [el('b', {text: b})]);
        drawTimeline(ca, r.keep); drawTimeline(cb, r.lose); c2.appendChild(ca); c2.appendChild(cb); det.appendChild(c2);
      } else if (kind == 'login') {
        function logins(list){ return list.map(function(u){ return item(u.user + ', made ' + u.created + (u.last && u.last != u.created ? ', last used ' + u.last : ', last used the day it was made') + (u.locked ? ', locked' : '')); }); }
        det.appendChild(el('div', {'class': 'cols'}, [col(a, logins(r.keep)), col(b, logins(r.lose))]));
        if (r.keep.length && r.lose.length) det.appendChild(el('div', {'class': 'muted', text: 'Both have a login. After the merge, ' + k.name + ' has both. Keep the one used most recently; an unused one can be removed on the person\'s User Account tab in TouchPoint.'}));
      }
    });
  }

  // a record's timeline, oldest first; around a name change the rows are bold
  function drawTimeline(box, list, focusDay){
    if (!list || !list.length) { box.appendChild(el('span', {'class': 'muted', text: 'Nothing recorded.'})); return; }
    list.forEach(function(e){
      var near = focusDay && e.at.slice(0, 10) == focusDay;
      box.appendChild(el('div', {'class': 'item', style: (e.kind == 'ident' ? 'color:#9B2C2C;font-weight:700' : (near ? 'font-weight:700' : ''))}, [el('span', {'class': 'muted', text: e.at + '  '}), e.text]));
    });
  }
  function localDetail(kind, det){
    var k = CUR.data.keep, l = CUR.data.lose, a = '#' + k.pid + ' (stays)', b = '#' + l.pid; det.dataset.loaded = '1'; det.innerHTML = '';
    function col(title, kids){ return el('div', {'class': 'stack', style: 'gap:2px'}, [el('b', {text: title})].concat(kids && kids.length ? kids : [el('span', {'class': 'muted', text: 'None'})])); }
    if (kind == 'evs') {
      var V = CUR.data.values, conf = {}, dup = {}; V.conflicts.forEach(function(c){ conf[c.field] = 1; }); V.dups.forEach(function(f){ dup[f] = 1; });
      function evs(list){ return list.map(function(x){ var tag = conf[x.field] ? ' (differs)' : (dup[x.field] ? ' (same on both)' : '');
        return el('div', {'class': 'item' + (conf[x.field] ? ' shared' : '')}, [el('b', {text: x.field + ': '}), (x.value || '(blank)') + tag]); }); }
      var vp = []; if (V.moves.length) vp.push(V.moves.length + ' move over as they are'); if (V.dups.length) vp.push(V.dups.length + ' are the same on both'); if (V.conflicts.length) vp.push(V.conflicts.length + ' differ (choose below)');
      det.appendChild(el('div', {'class': 'muted', text: vp.length ? vp.join(', ') + '.' : 'Nothing to move: #' + CUR.data.lose.pid + ' has no extra values of its own.'}));
      det.appendChild(el('div', {'class': 'cols'}, [col(a, evs(V.keep)), col(b, evs(V.lose))]));
    } else {
      var MD = CUR.data.medical;
      if (!MD.length) { det.appendChild(el('span', {'class': 'muted', text: 'Neither record has medical or emergency info.'})); return; }
      var words = {same: 'same on both', replaced: 'TouchPoint replaces it with #' + l.pid + "'s", filled: 'filled in from #' + l.pid, kept: 'stays'};
      var t2 = el('table', {'class': 'fields'}, [el('tr', {}, [el('th', {text: ''}), el('th', {text: a}), el('th', {text: b}), el('th', {text: 'After the merge'})])]);
      MD.forEach(function(m){ t2.appendChild(el('tr', {}, [el('td', {}, [el('b', {text: m.label})]), el('td', {text: m.keep || '(blank)'}), el('td', {text: m.other || '(blank)'}),
        el('td', {'class': m.what == 'replaced' ? 'bad' : '', text: words[m.what]})])); });
      det.appendChild(el('div', {'class': 'scroll'}, [t2]));
    }
  }

  // TouchPoint's own merge screen for this pair: every field it will post, and the
  // rows a person picks (keep this one's value or the other's)
  function getMergeForm(cb){
    var x = new XMLHttpRequest(); x.open('GET', '/Merge/' + CUR.lose + '/' + CUR.keep, true);
    x.onload = function(){
      var doc = new DOMParser().parseFromString(x.responseText, 'text/html');
      var form = null; Array.prototype.forEach.call(doc.querySelectorAll('form'), function(f){ if (!form && f.querySelector('[name="pi[0].PeopleId"]')) form = f; });
      if (!form) { cb("TouchPoint's merge screen did not open for this pair (check your role, or that both people still exist)."); return; }
      var p0 = form.querySelector('[name="pi[0].PeopleId"]').getAttribute('value'), p1 = form.querySelector('[name="pi[1].PeopleId"]').getAttribute('value');
      if (String(p0) != String(CUR.lose) || String(p1) != String(CUR.keep)) { cb("TouchPoint's merge screen opened a different pair. Nothing was changed."); return; }
      cb(null, form);
    };
    x.onerror = function(){ cb('Could not reach TouchPoint.'); };
    x.send();
  }
  function txt(n){ return n ? String(n.textContent || '').replace(/\s+/g, ' ').trim() : ''; }
  function placeholder(v){ v = String(v || '').trim();
    return /^(just added|unknown|n\/a|\(?not specified\)?|\(none\))$/i.test(v) || /^added during online registration because record was not found\.?$/i.test(v); }
  // two values one or two letters apart (Franklyn and Frankyln): a typo on one side, so nothing is picked for you
  function nearly(a, b){ a = String(a || '').toLowerCase(); b = String(b || '').toLowerCase(); if (!a || !b || a == b || Math.abs(a.length - b.length) > 2 || a.length < 5) return false;
    var d = [], i, j; for (i = 0; i <= a.length; i++) { d[i] = [i]; } for (j = 1; j <= b.length; j++) d[0][j] = j;
    for (i = 1; i <= a.length; i++) for (j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i-1][j] + 1, d[i][j-1] + 1, d[i-1][j-1] + (a[i-1] == b[j-1] ? 0 : 1));
      if (i > 1 && j > 1 && a[i-1] == b[j-2] && a[i-2] == b[j-1]) d[i][j] = Math.min(d[i][j], d[i-2][j-2] + 1); }
    return d[a.length][b.length] <= 2; }
  // the other value is the same one, only more complete (1/9 and 1/9/1952, 37075 and 37075-1733)
  function fuller(keep, other){ keep = String(keep || '').trim().toLowerCase(); other = String(other || '').trim().toLowerCase();
    return !!keep && other.length > keep.length && other.indexOf(keep) == 0; }
  // one last name is a part of the other (Ge Matos and Ge, Vitier-Farmer and Farmer): the same name, not a change
  function partOf(a, b){ a = String(a || '').toLowerCase().trim(); b = String(b || '').toLowerCase().trim(); if (!a || !b || a == b) return false;
    return a.split(/[\s-]+/).indexOf(b) >= 0 || b.split(/[\s-]+/).indexOf(a) >= 0; }
  function digitsOff(a, b){ a = String(a || '').replace(/\D/g, ''); b = String(b || '').replace(/\D/g, ''); if (!a || !b || a == b || a.length != b.length) return false;
    var d = [], i; for (i = 0; i < a.length; i++) if (a.charAt(i) != b.charAt(i)) d.push(i);
    return d.length == 1 || (d.length == 2 && d[1] == d[0] + 1 && a.charAt(d[0]) == b.charAt(d[1]) && a.charAt(d[1]) == b.charAt(d[0])); }
  function loadFields(){
    getMergeForm(function(err, form){
      var box = $('mgfields'); if (!box) return;
      box.innerHTML = ''; box.appendChild(el('h2', {text: 'Where they differ'}));
      if (err) { box.appendChild(el('span', {'class': 'bad', text: err})); return; }
      var seen = {}; FIELDS = [];
      var LOSE_NEWER = (CUR.data.lose.created || '') > (CUR.data.keep.created || '');
      var BOTH_ALONE = CUR.data.lose.familySize <= 1 && CUR.data.keep.familySize <= 1;
      // Married into the family being joined: #keep moves into #lose's family, #lose has a different last
      // name (not a typo of it), and someone else in that family has #lose's last name. Then #lose's last
      // name, title and marital status are the current ones, and #keep's last name becomes the former name.
      function hv(n){ var e = form.querySelector('[name="' + n + '"]'); return e ? String(e.getAttribute('value') || '').trim() : ''; }
      var LAST_K = hv('pi[1].LastName'), LAST_L = hv('pi[0].LastName');
      var famLast = LAST_L && (CUR.data.loseFamily || []).some(function(m){ var n = String(m.name || '').toLowerCase(); return !m.self && n.slice(-LAST_L.length) == LAST_L.toLowerCase(); });
      var NICK = CUR.data.nick || null;
      // last names used by the rest of each family, and how common each spelling is
      function famNames(list){ return (list || []).filter(function(m){ return !m.self && m.last; }).map(function(m){ return m.last; }); }
      var FAMK = famNames(CUR.data.keepFamily), FAML = famNames(CUR.data.loseFamily);
      function inFam(list, v){ v = String(v || '').toLowerCase(); return !!v && list.some(function(n){ return n.toLowerCase() == v; }); }
      var SP = CUR.data.spell || {first: {}, last: {}};
      function usual(kind, a, b){ var ca = SP[kind][String(a).toLowerCase()] || 0, cb = SP[kind][String(b).toLowerCase()] || 0, mn = Math.min(ca, cb), mx = Math.max(ca, cb);
        return (mn == 0 && mx >= 1) || (mn <= 1 && mx >= 5) ? (ca > cb ? a : b) : ''; }
      famLast = famLast || inFam(FAML, LAST_L);
      var MARRIED_IN = !!(CUR.joinFamily && LAST_K && LAST_L && LAST_K.toLowerCase() != LAST_L.toLowerCase() && !nearly(LAST_K, LAST_L) && famLast);
      Array.prototype.forEach.call(form.querySelectorAll('input[type=radio]'), function(r0){
        var name = r0.getAttribute('name') || ''; if (!/^Use/.test(name) || seen[name] || r0.getAttribute('value') != '0') return;
        var tr = r0.closest('tr'); if (!tr) return;
        var r1 = tr.querySelector('input[type=radio][name="' + name + '"][value="1"]'); if (!r1) return;
        seen[name] = 1;
        var c0 = r0.closest('td'), c1 = r1.closest('td');
        var label = txt(tr.querySelector('td')), v0 = txt(c0 && c0.nextElementSibling), v1 = txt(c1 && c1.nextElementSibling);
        // TouchPoint's merge screen labels grade level, school grade and shirt size all "Grade": use the field's real name
        if (name == 'UseGradeLevelId') label = 'Grade level'; else if (name == 'UseGrade') label = 'School grade'; else if (name == 'UseShirtSizeId') label = 'Shirt size';
        var def = r0.hasAttribute('checked') ? '0' : '1';
        var family = /^(address line|city|state code|zip|country|home phone)/i.test(label);
        var contact = !family && /cell|email|work phone/i.test(label);
        var differs = v0 && v0.toLowerCase() != v1.toLowerCase();
        // keep the kept person's value, but fill in from the other when the kept one is blank, and
        // take a personal phone or email from the newer record (just typed, so most likely current)
        // a family's address and home phone never fill in from a different family, unless both records
        // are alone in their families (then those values are that one person's own)
        var crossFamily = family && !CUR.data.sameFamily && !BOTH_ALONE;
        // a value TouchPoint or a kiosk fills in by itself (Just Added, Unknown, not specified) is not
        // information: a real value on the other record wins over it
        var ph1 = placeholder(v1), ph0 = placeholder(v0);
        // a phone typo is one digit wrong or two swapped; any more and it is a different number
        var typo = /phone/i.test(label) ? digitsOff(v1, v0) : nearly(v1, v0), more = fuller(v1, v0) && !crossFamily;
        var pick = def == '0' ? '0' : (crossFamily ? '1' : ((!v1 || ph1) && v0 && !ph0 ? '0' : (more ? '0' : (LOSE_NEWER && contact && differs && !typo ? '0' : '1'))));
        // the membership dates (join, decision, baptism, drop) only come along with the member status
        var memField = /^member status/i.test(label);
        // how a record was made describes that record: the duplicate's origin or entry point is never the person's
        if (/^(origin|entry point|interest point)/i.test(label) && def != '0') pick = '1';
        if (memField && differs && CUR.data.lose.membership.length && !CUR.data.keep.membership.length) pick = '0';
        // TouchPoint's own third column: a value typed in (text fields only)
        var cin = tr.querySelector('input[type=text][name^="pi[2]."]'), customName = cin ? cin.getAttribute('name') : '', custom = '';
        var married = MARRIED_IN && /^(last name|title|marital status)/i.test(label) && differs && !ph0;
        if (married) pick = '0';
        var formerSet = false;
        // Judith and Judy: the first name is the formal one, the short one goes in Goes By (Nick Name)
        if (NICK && /^first name/i.test(label) && NICK.formalSide == 'lose') pick = '0';
        var goesBy = !!(NICK && /^nick name/i.test(label) && !v0 && !v1 && customName);
        if (goesBy) { pick = '2'; custom = NICK.short; }
        // a misspelled name: the spelling the family uses, else the usual spelling when the other is rare
        var isLast = /^last name/i.test(label), isFirst = /^first name/i.test(label), note = null, newerName = false;
        if ((isLast || (isFirst && !NICK)) && differs && typo && !married) {
          var both = FAMK.concat(FAML), fh = isLast && inFam(both, v0) ? v0 : (isLast && inFam(both, v1) ? v1 : '');
          var win = fh || usual(isLast ? 'last' : 'first', v1, v0);
          if (win) { pick = win.toLowerCase() == v0.toLowerCase() ? '0' : '1'; note = {v: pick, text: fh ? 'the family\'s spelling' : 'the usual spelling'}; }
        }
        // a different last name (not a typo, not one part of the other): keep the one the family shares;
        // otherwise the newer record's is most likely current (a marriage), unless it is a misspelling
        // of the family's own name (Grenior for Grenier), then the family's spelling is typed in
        // an email that bounced, was blocked or reported spam in the last year is not picked over one that works
        var EF = CUR.data.emailFails || {};
        if (/^email address/i.test(label) && differs && v0 && v1) {
          var e0 = EF[v0.toLowerCase()], e1 = EF[v1.toLowerCase()];
          if (e0 && e0.recent && !e1) { pick = '1'; note = {v: '1', text: 'works: the other ' + e0.text}; }
          else if (e1 && e1.recent && !e0) { pick = '0'; note = {v: '0', text: 'works: the other ' + e1.text}; }
        }
        if (isLast && differs && !typo && !married && !ph0 && !partOf(v0, v1)) {
          if (!CUR.joinFamily && inFam(FAMK, v1)) note = {v: '1', text: 'the family\'s last name'};
          else if (LOSE_NEWER) {
            var famSp = FAMK.concat(CUR.joinFamily ? FAML : []).filter(function(n){ return nearly(n, v0) && n.toLowerCase() != v0.toLowerCase(); })[0];
            if (famSp && customName) { pick = '2'; custom = famSp; note = {v: '2', text: famSp + ': the family\'s spelling of ' + v0}; }
            else { pick = '0'; newerName = true; note = {v: '0', text: 'newer: probably a name change'}; }
          }
        }
        FIELDS.push({name: name, label: label, keep: v1, other: v0, pick: pick, def: def, family: family, newer: LOSE_NEWER && differs && (contact || family), ph1: ph1 && !!v1, ph0: ph0 && !!v0, mem: memField, typo: typo, more: more,
          customName: customName, custom: custom, married: married, former: formerSet, goesBy: goesBy, note: note, newerName: newerName,
          isLast: isLast, isFormer: /^former name/i.test(label)});
      });
      // Former Name follows the last name: when the last name kept drops a different one (a marriage,
      // a name change), the dropped one is typed into Former Name, unless a record already has one
      var LF = FIELDS.filter(function(f){ return f.isLast; })[0], FF = FIELDS.filter(function(f){ return f.isFormer; })[0];
      function syncFormer(){
        if (!LF || !FF || !FF.customName || FF.touched) return;
        if ((FF.keep && !FF.ph1) || (FF.other && !FF.ph0)) return;
        var fin = LF.pick == '0' ? LF.other : (LF.pick == '2' ? LF.custom : LF.keep);
        // only the older record's last name can be a former name: the newer one may be the current name
        var dropped = [LF.keep, LF.other].filter(function(v, i){ return v && fin && v.toLowerCase() != fin.toLowerCase() && !nearly(v, fin) && !partOf(v, fin) && (i == 0 ? LOSE_NEWER : !LOSE_NEWER); })[0] || '';
        if (dropped) { FF.pick = '2'; FF.custom = dropped; FF.former = true;
          FF.formerWhy = MARRIED_IN && LF.pick == '0' ? '#' + CUR.data.keep.pid + '\'s last name before joining the ' + LAST_L + ' family' : 'The other last name on these records'; }
        else if (FF.former) { FF.pick = '1'; FF.custom = ''; FF.former = false; }
      }
      function refreshFormer(){ syncFormer(); if (!FF || !FF.ui || !FF.ui.typed) return; var u = FF.ui;
        u.tin.value = FF.custom || ''; u.tr2.checked = FF.pick == '2'; if (u.r1) u.r1.checked = FF.pick == '1'; if (u.r0) u.r0.checked = FF.pick == '0';
        u.typed.hidden = FF.pick != '2'; u.lbl.textContent = FF.former ? FF.formerWhy : 'Typed value';
        var sm = FF.keep.toLowerCase() == FF.other.toLowerCase() && FF.pick != '2'; u.row.className = sm ? 'same' : ''; u.typed.className = sm ? 'same' : ''; }
      syncFormer();
      var diff = FIELDS.filter(function(f){ return f.keep.toLowerCase() != f.other.toLowerCase(); });
      var table = el('table', {'class': 'fields'}, [el('tr', {}, [el('th', {text: 'Field'}), el('th', {text: 'Keep ' + CUR.data.keep.name + "'s"}), el('th', {text: 'Use ' + CUR.data.lose.name + "'s"})])]);
      FIELDS.forEach(function(f){
        var same = f.keep.toLowerCase() == f.other.toLowerCase();
        function opt(v, value){ var i = el('input', {type: 'radio', name: 'mg_' + f.name, value: v}); if (f.pick == v) i.checked = true;
          i.onchange = function(){ f.pick = v; if (f.isFormer) f.touched = true; if (f.isLast) refreshFormer(); };
          (f.ui = f.ui || {})['r' + v] = i;
          var ph = v == '1' ? f.ph1 : f.ph0;
          return el('label', {}, [i, el('span', {text: value || '(blank)'}), v == '0' && f.newer && !(f.family && CUR.joinFamily) ? el('span', {'class': 'tag', style: 'margin-left:6px', text: 'newer'}) : null,
            v == '0' && f.typo && !f.note ? el('span', {'class': 'tag warn', style: 'margin-left:6px', text: /phone/i.test(f.label) ? 'a digit off: check which is right' : 'nearly the same: check the spelling'}) : null,
            f.note && f.note.v == v ? el('span', {'class': 'tag', style: 'margin-left:6px', text: f.note.text}) : null,
            (function(){ var fl = /^email address/i.test(f.label) && value ? (CUR.data.emailFails || {})[String(value).toLowerCase()] : null;
              return fl ? el('span', {'class': 'tag warn', style: 'margin-left:6px', text: fl.text + ', last ' + fl.last}) : null; })(),
            v == '0' && f.more ? el('span', {'class': 'tag', style: 'margin-left:6px', text: 'more complete'}) : null,
            v == '0' && f.married ? el('span', {'class': 'tag', style: 'margin-left:6px', text: 'current: the family it joins'}) : null,
            ph ? el('span', {'class': 'muted', style: 'margin-left:6px;font-size:12px', text: 'a default, not information'}) : null]); }
        var mainRow = el('tr', {'class': same && f.pick != '2' ? 'same' : ''}, [el('td', {}, [el('b', {text: f.label}), f.family ? el('div', {'class': 'muted', style: 'font-size:12px', text: CUR.joinFamily ? 'Not used: #' + CUR.data.keep.pid + ' takes the ' + (LAST_L ? LAST_L + ' ' : 'other ') + 'family\'s address and home phone' : 'Belongs to the whole family'}) : null,
            f.mem ? el('div', {'class': 'muted', style: 'font-size:12px', text: 'Brings that record\'s join, decision, baptism and drop dates with it'}) : null,
            CUR.joinFamily && /^position in family/i.test(f.label) ? el('div', {'class': 'muted', style: 'font-size:12px', text: 'Set by the family #' + CUR.data.keep.pid + ' joins (adult or child by age)'}) : null,
            /^(origin|entry point|interest point)/i.test(f.label) ? el('div', {'class': 'muted', style: 'font-size:12px', text: 'How each record was made; the original\'s stays'}) : null,
            f.customName && f.pick != '2' && (!same || /^former name/i.test(f.label)) && !(f.family && CUR.joinFamily) ? (function(){ var lb = el('button', {type: 'button', 'class': 'linkbtn', style: 'display:block;font-size:12px;margin-top:2px;padding:0', text: 'Type another value'});
              lb.onclick = function(){ lb.hidden = true; typed.hidden = false; tin.focus(); }; return lb; })() : null]),
          el('td', {}, [opt('1', f.keep)]), el('td', {}, [opt('0', f.other)])]);
        table.appendChild(mainRow);
        // a value neither record has (fixing a typo, or a former name), using TouchPoint's own typed-value column
        if (f.customName) {
          var tr2 = el('input', {type: 'radio', name: 'mg_' + f.name, value: '2'}); if (f.pick == '2') tr2.checked = true;
          var tin = el('input', {type: 'text', value: f.custom || '', style: 'max-width:260px', 'aria-label': f.label + ', typed value'});
          tr2.onchange = function(){ f.pick = '2'; if (f.isFormer) f.touched = true; if (f.isLast) refreshFormer(); };
          tin.oninput = function(){ f.custom = tin.value; f.pick = '2'; tr2.checked = true; if (f.isFormer) f.touched = true; if (f.isLast) refreshFormer(); };
          var lbl = el('span', {'class': 'muted', style: 'font-size:12px', text: f.former ? f.formerWhy : (f.goesBy ? 'Goes by ' + NICK.short + ' (the first name on #' + (NICK.formalSide == 'keep' ? CUR.data.lose.pid : CUR.data.keep.pid) + ')' : (f.note && f.note.v == '2' ? f.note.text : 'Typed value'))});
          var typed = el('tr', {'class': same && f.pick != '2' ? 'same' : ''}, [el('td', {}, [lbl]), el('td', {colspan: '2'}, [el('label', {}, [tr2, tin])])]);
          typed.hidden = f.pick != '2';
          table.appendChild(typed);
          f.ui = f.ui || {}; f.ui.tin = tin; f.ui.tr2 = tr2; f.ui.typed = typed; f.ui.lbl = lbl; f.ui.row = mainRow;
        }
      });
      var k = CUR.data.keep, l = CUR.data.lose;
      box.appendChild(el('div', {'class': 'muted', text: diff.length ? plural(diff.length, 'field differs', 'fields differ') + '. Blank values on #' + k.pid + ' are filled in from #' + l.pid + '.' +
        (LOSE_NEWER ? ' Cell phone and email default to the newer record (#' + l.pid + ', added ' + l.created + ').' : '') +
        (CUR.data.sameFamily || CUR.joinFamily ? '' : (BOTH_ALONE ? ' Each is alone in their family, so address and home phone are treated like their own: blanks fill in from the other record.' : ' Address and home phone belong to the whole family, so they stay as they are unless you choose otherwise.')) +
        ' Values TouchPoint or a kiosk fills in by itself (Just Added, Unknown, not specified) give way to a real value on the other record.' +
        (NICK ? ' ' + NICK.short + ' is a short form of ' + NICK.formal + ', so the first name is ' + NICK.formal + ' and ' + NICK.short + ' goes in Nick Name (Goes By), which TouchPoint shows as the name.' : '') +
        (MARRIED_IN ? ' #' + k.pid + ' is joining the ' + LAST_L + ' family, so last name, title and marital status come from #' + l.pid + ', and ' + LAST_K + ' is saved as the former name.' : '') +
        (LF && LF.newerName ? ' The last names differ and #' + l.pid + ' is newer, so ' + LF.other + ' is taken as the current last name' + (FF && FF.former ? ' and ' + FF.custom + ' is saved as the former name' : '') + ': choose the other if that is wrong.' : '') +
        (LF && LF.note && LF.note.v == '2' ? ' ' + LF.note.text + ' is typed in as the last name.' : '') : 'Every field is the same on both.'}));
      box.appendChild(el('div', {'class': 'scroll'}, [table]));
      var all = el('button', {type: 'button', 'class': 'btn ghost', text: 'Show every field'}); all.onclick = function(){ table.classList.toggle('all'); all.textContent = table.classList.contains('all') ? 'Show only differences' : 'Show every field'; };
      box.appendChild(el('div', {}, [all]));
      var nx = $('mgnext'); if (nx) nx.disabled = false;
    });
  }

  // ---- before you merge
  function preview(){
    var r = CUR.data, k = r.keep, l = r.lose, lc = l.counts, loss = r.loss;
    show('mgpreview');
    $('mgpsub').textContent = l.name + ' (#' + l.pid + ') goes into ' + k.name + ' (#' + k.pid + '), then #' + l.pid + ' is deleted.';
    var moves = [];
    if (lc.gifts == null) moves.push('every gift, keeping its date and fund (how many shows to the Finance role)');
    else if (lc.gifts) moves.push(plural(lc.gifts, 'gift', 'gifts') + ', keeping their dates and funds');
    if (lc.att) moves.push(plural(lc.att, 'attended meeting', 'attended meetings'));
    if (lc.orgs) moves.push(plural(lc.orgs, 'involvement', 'involvements'));
    if (lc.regs) moves.push(plural(lc.regs, 'registration', 'registrations'));
    if (lc.logins) moves.push('their login');
    if (lc.checks) moves.push(plural(lc.checks, 'background check', 'background checks'));
    if (lc.notes) moves.push(plural(lc.notes, 'note or task', 'notes and tasks'));
    if (lc.flags) moves.push(plural(lc.flags, 'status flag', 'status flags'));
    moves.push('emails, texts, opt-outs, volunteer records and the rest of their history');
    var chosen = FIELDS.filter(function(f){ return f.pick == '0' && f.keep.toLowerCase() != f.other.toLowerCase(); });
    chosen.forEach(function(f){ moves.push(f.label + ': ' + (f.other || '(blank)') + (f.family ? ' (for the whole family)' : '')); });
    FIELDS.filter(function(f){ return f.pick == '2'; }).forEach(function(f){ moves.push(f.label + ': ' + (f.custom || '(blank)') + ' (typed in)'); });
    var saves = [];
    var addTags = loss.tags.filter(function(t){ return !t.has; });
    if (addTags.length) saves.push('Add ' + k.name + ' (#' + k.pid + ') to ' + plural(addTags.length, 'tag', 'tags') + ' #' + l.pid + ' was in: ' + addTags.map(function(t){ return t.name; }).join(', '));
    loss.ownedTags.forEach(function(t){ saves.push('Rebuild the tag ' + t.name + ' (' + plural(t.n, 'person', 'people') + '), owned by ' + k.name); });
    var cy = loss.carry || {};
    if (cy.checkInNote) saves.push(cy.checkInBoth ? 'Add #' + l.pid + '\'s check-in note to #' + k.pid + '\'s (both kept). TouchPoint\'s merge leaves the check-in note behind.' : 'Carry over #' + l.pid + '\'s check-in note (shown at children\'s check-in). TouchPoint\'s merge leaves it behind.');
    if (cy.msUserId) saves.push('Keep #' + l.pid + '\'s MinistrySafe account link, so its training and checks stay with ' + k.name + '. TouchPoint\'s merge leaves it behind, and MinistrySafe would get a second account.');
    var lostVals = loss.familyValues.filter(function(v){ return !v.kept; }), fv = loss.familyValues.map(function(v){ return v.field; }).join(', ');
    if (loss.valuesMode == 'same' && loss.familyValues.length) saves.push('Put back the family\'s ' + plural(loss.familyValues.length, 'extra value', 'extra values') + ' (' + fv + '). TouchPoint deletes a family\'s extra values when it deletes anyone in it.');
    else if (loss.valuesMode == 'their' && loss.familyValues.length) saves.push('Put back the ' + plural(loss.familyValues.length, 'extra value', 'extra values') + ' of the family with ' + names(loss.leavesFamily) + ' (' + fv + '). TouchPoint deletes a family\'s extra values when it deletes anyone in it.');
    else if (lostVals.length) saves.push('Copy ' + plural(lostVals.length, 'family extra value', 'family extra values') + ' to the family of #' + k.pid + ': ' + lostVals.map(function(v){ return v.field; }).join(', '));
    var V = CUR.data.values, picksOther = V.conflicts.filter(function(c){ return CUR.evPick[c.field] == 'other'; });
    if (V.conflicts.length || V.dups.length) saves.push('Extra values: ' + [V.conflicts.length ? plural(V.conflicts.length, 'field', 'fields') + ' settled as you chose' + (picksOther.length ? ' (' + picksOther.map(function(c){ return c.field; }).join(', ') + ' from #' + l.pid + ')' : '') : '',
      V.dups.length ? plural(V.dups.length, 'duplicate copy', 'duplicate copies') + ' removed' : ''].filter(Boolean).join('; ') + ', so no _mv leftovers');
    CUR.data.medical.filter(function(m){ return m.what == 'replaced'; }).forEach(function(m){ moves.push(m.label + ': "' + m.other + '" replaces "' + m.keep + '"'); });
    var gone = [];
    if (CUR.joinFamily) saves.push('Move ' + k.name + ' (#' + k.pid + ') into the family with ' + names(loss.leavesFamily) + ', with TouchPoint\'s Person > Family > Add');
    if (CUR.joinFamily && (r.keepRelated || []).length) gone.push('#' + k.pid + "'s related-family link" + (r.keepRelated.length == 1 ? '' : 's') + ' (' + r.keepRelated.map(function(x){ return x.name + (x.rel ? ', ' + x.rel : ''); }).join('; ') + ')' + (r.keepRelated.length == 1 ? ' stays' : ' stay') + ' on its old family, now empty. Add ' + (r.keepRelated.length == 1 ? 'it' : 'them') + ' again on the new family in TouchPoint.');
    if ((l.created || '') < (k.created || '')) gone.push('#' + l.pid + "'s record number and added date (" + l.created + '). #' + k.pid + ' keeps its own (' + k.created + '). Swap to keep the original instead.');
    var memPick = FIELDS.filter(function(f){ return f.mem; })[0];
    if (l.membership.length && (!memPick || memPick.pick != '0')) gone.push('#' + l.pid + "'s membership history (" + l.membership.join('; ') + '), since member status stays #' + k.pid + "'s. Choose #" + l.pid + "'s member status to keep it.");
    if (loss.familyGone) gone.push('#' + l.pid + "'s family of one, now empty" + (loss.relatedGone ? ', and its ' + plural(loss.relatedGone, 'related-family link', 'related-family links') : ''));
    if (loss.leavesFamily.length && !CUR.joinFamily) gone.push('#' + l.pid + "'s place in the family with " + names(loss.leavesFamily) + ': ' + k.name + ' stays in the family of #' + k.pid + '.');
    if (loss.headless) gone.push('#' + l.pid + ' is the head of ' + (loss.leavesFamily.length ? 'the family with ' + names(loss.leavesFamily) + '; afterward that family has no head of household. Add them to the right family in TouchPoint (Person > Family > Add).' : 'this family; afterward set the head of household on the family in TouchPoint.'));
    if (loss.absent) gone.push(plural(loss.absent, 'absent attendance row', 'absent attendance rows') + ' (TouchPoint moves only meetings attended)');
    if (loss.cards) gone.push(plural(loss.cards, 'saved check or card match', 'saved check or card matches') + ': the next gift from that account may need matching by hand');
    (loss.more || []).forEach(function(x){ gone.push(x); });
    if (l.photo && !chosen.some(function(f){ return /photo/i.test(f.label); })) gone.push('#' + l.pid + "'s photo");
    if (CUR.overwritten) gone.unshift('A record here was overwritten with a different person (see the red box on the side by side). Merging deletes the other person.');
    gone.push('There is no undo; TouchPoint deletes the record.');
    function card(title, color, items, foot, cls){ return el('div', {'class': 'card' + (cls ? ' ' + cls : '')}, [el('div', {'class': 'eyebrow', style: 'color:' + color, text: title}),
      items.length ? el('ul', {}, items.map(function(t){ return el('li', {text: t}); })) : el('span', {'class': 'muted', text: 'Nothing.'}), foot ? el('span', {'class': 'muted', style: 'font-size:13px', text: foot}) : null]); }
    var body = $('mgpbody'); body.innerHTML = '';
    body.appendChild(el('div', {'class': 'grid2', style: 'grid-template-columns:repeat(auto-fit,minmax(280px,1fr))'}, [
      card('TouchPoint moves into #' + k.pid, '#1F5A8A', moves, "Done by TouchPoint's own merge, as you, and recorded in its merge history."),
      card('Saved by this page', '#1E6B45', saves, saves.length ? "TouchPoint's merge would drop these." : null),
      card('Gone with #' + l.pid, '#7A4A00', gone, null, 'box')]));
    // the people left in the deleted record's family can come along into the kept family
    // (TouchPoint's own Person > Family > Add, done right after the merge)
    CUR.bring = []; CUR.bringNew = [];
    // moving the original into the other family: its own family can come along (Myles, with his mother)
    if (CUR.joinFamily && loss.keepOthers.length) {
      var onlyKids = loss.keepOthers.every(function(x){ return x.age != null && x.age < 18; });
      var bn = el('div', {'class': 'stack', style: 'gap:6px'});
      loss.keepOthers.forEach(function(x){
        var on = onlyKids;   // children who would otherwise be left with no adult
        var c = el('input', {type: 'checkbox', style: 'width:20px;height:20px'}); c.checked = on; if (on) CUR.bringNew.push(x);
        c.onchange = function(){ CUR.bringNew = CUR.bringNew.filter(function(y){ return y.pid != x.pid; }); if (c.checked) CUR.bringNew.push(x); };
        bn.appendChild(el('label', {'class': 'row', style: 'font-weight:600'}, [c, x.name + (x.age != null ? ' (' + x.age + ')' : '')]));
      });
      body.appendChild(el('div', {'class': 'card'}, [el('h2', {text: 'Also bring into the family with ' + names(loss.leavesFamily) + '?'}),
        el('div', {'class': 'muted', text: 'They are in the family ' + k.name + ' leaves. Checked people move with ' + k.name + ' right after the merge (Person > Family > Add in TouchPoint).' + (onlyKids ? ' The children are checked so no one is left without an adult.' : '')}), bn]));
    }
    if (loss.leavesFamily.length && !CUR.joinFamily) {
      var bl = el('div', {'class': 'stack', style: 'gap:6px'});
      loss.leavesFamily.forEach(function(x){
        var on = loss.headless && x.age != null && x.age < 18;   // a child who would be left with no adult
        var c = el('input', {type: 'checkbox', style: 'width:20px;height:20px'}); c.checked = on; if (on) CUR.bring.push(x);
        c.onchange = function(){ CUR.bring = CUR.bring.filter(function(y){ return y.pid != x.pid; }); if (c.checked) CUR.bring.push(x); };
        bl.appendChild(el('label', {'class': 'row', style: 'font-weight:600'}, [c, x.name + (x.age != null ? ' (' + x.age + ')' : '')]));
      });
      body.appendChild(el('div', {'class': 'card'}, [el('h2', {text: 'Also bring into the family of #' + k.pid + '?'}),
        el('div', {'class': 'muted', text: 'They are in the family of #' + l.pid + ', which ' + k.name + ' leaves. Checked people are moved into the family of #' + k.pid + ' right after the merge, the same as Person > Family > Add in TouchPoint.' +
          (loss.headless ? ' Anyone left behind is in a family with no head of household.' : '')}), bl]));
    }
    var ok = el('input', {type: 'checkbox', style: 'width:20px;height:20px'});
    var go = el('button', {type: 'button', 'class': 'btn warn', text: 'Merge', disabled: 'disabled'});
    ok.onchange = function(){ go.disabled = !ok.checked; };
    go.onclick = function(){ go.disabled = true; merge(); };
    body.appendChild(el('div', {'class': 'card'}, [el('label', {'class': 'row', style: 'font-weight:600'}, [ok, 'I understand ' + l.name + ' (#' + l.pid + ') will be deleted']),
      el('div', {'class': 'row'}, [el('span', {'class': 'muted', style: 'flex:1;min-width:240px', text: 'TouchPoint saves a merge step by step, so afterward this page checks that everything landed on #' + k.pid + '.'}),
        el('button', {type: 'button', 'class': 'btn ghost', text: 'Back', onclick: function(){ show('mgreview'); }}), go])]));
  }
  $('mgback2').onclick = function(){ show('mgreview'); };

  // ---- the merge
  function logLine(cls, text){ var d = el('div', {'class': cls || '', text: text}); $('mgdlog').appendChild(d); return d; }
  function merge(){
    var k = CUR.data.keep, l = CUR.data.lose;
    show('mgdone'); $('mgdtitle').textContent = 'Merging ' + l.name + ' into ' + k.name + '...'; $('mgdlog').innerHTML = ''; $('mgdbtns').hidden = true;
    $('mgdopen').href = '/Person2/' + k.pid;
    // the merge log: the before step starts it; finishing adds what happened (never giving)
    var LOGID = '';
    function finish(title){ $('mgdtitle').textContent = title; $('mgdbtns').hidden = false;
      if (!LOGID) return;
      var lines = []; Array.prototype.forEach.call($('mgdlog').children, function(d){ if (/^(ok|bad|muted)$/.test(d.className)) lines.push({cls: d.className, text: d.textContent}); });
      post({action: 'log', keep: k.pid, id: LOGID, lines: JSON.stringify(lines)}, function(){}); LOGID = ''; }
    function pickSummary(){ return FIELDS.filter(function(f){ return (f.pick != f.def || f.pick == '2') && !/envelope|statement|contribut|giving/i.test(f.label); })
      .map(function(f){ return {field: f.label, chose: f.pick == '0' ? 'from the deleted record' : (f.pick == '2' ? 'typed' : 'kept'),
        value: String((f.pick == '0' ? f.other : (f.pick == '2' ? f.custom : f.keep)) || '').slice(0, 200)}; }); }
    // 1. a fresh copy of TouchPoint's merge form, before anything changes
    getMergeForm(function(err, form){
      if (err) { logLine('bad', err); finish('Not merged'); return; }
      var pairs = [];
      Array.prototype.forEach.call(form.elements, function(e){
        var n = e.getAttribute('name'); if (!n || e.disabled) return;
        var t = (e.getAttribute('type') || '').toLowerCase(), tag = e.tagName.toLowerCase();
        if (t == 'radio' || t == 'checkbox') { if (e.hasAttribute('checked')) pairs.push([n, e.getAttribute('value') || 'on']); return; }
        if (t == 'submit' || t == 'button' || t == 'file') return;
        if (tag == 'select') { var o = e.querySelector('option[selected]') || e.querySelector('option'); pairs.push([n, o ? (o.hasAttribute('value') ? o.getAttribute('value') : o.textContent) : '']); return; }
        if (tag == 'textarea') { pairs.push([n, e.textContent]); return; }
        pairs.push([n, e.getAttribute('value') || '']);
      });
      // the choices made here, and delete the duplicate (Merge Fields and Move Related, with Delete checked)
      var choose = {}, typedVal = {}; FIELDS.forEach(function(f){ choose[f.name] = f.pick; if (f.pick == '2' && f.customName) typedVal[f.customName] = f.custom || ''; });
      pairs = pairs.filter(function(p){ return !(p[0] in choose) && !(p[0] in typedVal) && p[0] != 'DeleteDuplicate'; });
      for (var nm in choose) pairs.push([nm, choose[nm]]);
      for (var tn in typedVal) pairs.push([tn, typedVal[tn]]);
      pairs.push(['DeleteDuplicate', 'true']);
      // 2. what this page saves first
      post({action: 'before', keep: k.pid, lose: l.pid, picks: JSON.stringify(pickSummary()),
            move: CUR.joinFamily ? 'moved into the family with ' + names(CUR.data.loss.leavesFamily) : ''}, function(b){
        if (!b.success) { logLine('bad', 'Stopped before merging: ' + b.message); finish('Not merged'); return; }
        LOGID = b.logId || '';
        b.log.forEach(function(t){ logLine(/could not/.test(t) ? 'bad' : 'ok', t); });
        if (!b.log.length) logLine('muted', 'No tags needed saving.');
        // 3. TouchPoint's own merge
        var m = logLine('busy', "TouchPoint is merging (this can take a minute)..."); m.insertBefore(el('span', {'class': 'spin'}), m.firstChild);
        var x = new XMLHttpRequest(); x.open('POST', '/Merge/MergeFieldsAndMoveRelated', true);
        x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        x.onload = function(){
          var landed = /\/Person2\//i.test(x.responseURL || '');
          m.className = landed ? 'ok' : 'bad';
          m.textContent = landed ? 'TouchPoint merged ' + l.name + ' (#' + l.pid + ') into ' + k.name + ' (#' + k.pid + ').' : 'TouchPoint did not finish the way it usually does (status ' + x.status + '). Checking what happened...';
          after();
        };
        x.onerror = function(){ m.className = 'bad'; m.textContent = 'Lost touch with TouchPoint during the merge. Checking what happened...'; after(); };
        x.send(pairs.map(function(p){ return encodeURIComponent(p[0]) + '=' + encodeURIComponent(p[1]); }).join('&'));
        // 4. family values back, 5. check it all landed
        function after(){
          post({action: 'after', keep: k.pid, anchor: b.anchor || 0, familyValues: JSON.stringify(b.familyValues || []), carry: JSON.stringify(b.carry || {})}, function(a){
            if (a.success) a.log.forEach(function(t){ logLine(/could not/.test(t) ? 'bad' : 'ok', t); }); else logLine('bad', a.message);
            var picks = {}; for (var f in (CUR.evPick || {})) if (CUR.evPick[f] != 'both') picks[f] = CUR.evPick[f];
            post({action: 'values', keep: k.pid, picks: JSON.stringify(picks), dups: JSON.stringify(CUR.data.values.dups), keepFields: JSON.stringify(b.keepFields || [])}, function(ev){
            if (ev.success) ev.log.forEach(function(t){ logLine(/could not/.test(t) ? 'bad' : 'ok', t); }); else logLine('bad', ev.message);
            var repl = CUR.data.medical.filter(function(m){ return m.what == 'replaced'; });
            if (repl.length) { var md = logLine('bad', 'Check medical and emergency info on #' + k.pid + ': TouchPoint now has ' + repl.map(function(m){ return m.label + ' "' + m.other + '"'; }).join(', ') + '. If the old values were the current ones, change them back: ');
              md.appendChild(el('a', {href: '/Person2/' + k.pid, target: '_blank', rel: 'noopener', text: 'open #' + k.pid})); }
            post({action: 'verify', keep: k.pid, lose: l.pid, expect: JSON.stringify(b.expect || {})}, function(v){
              if (!v.success) { logLine('bad', v.message); finish('Check this pair in TouchPoint'); return; }
              logLine(v.gone ? 'ok' : 'bad', v.gone ? 'The duplicate, ' + l.name + ' (#' + l.pid + '), is gone.' : 'The duplicate, ' + l.name + ' (#' + l.pid + '), is still in TouchPoint.');
              if (v.name && v.name != k.name) { logLine('ok', k.name + ' (#' + k.pid + ') is now ' + v.name + (v.former ? ', with ' + v.former + ' as the former name' : '') + '.'); k.name = v.name; }
              var short = v.items.filter(function(i){ return !i.ok; });
              v.items.forEach(function(i){ if (i.want) logLine(i.ok ? 'ok' : 'bad', k.name + ' (#' + k.pid + ') now has ' + i.got + ' ' + i.what + (i.ok ? ', everything from both records' : ' (expected ' + i.want + ')')); });
              joinFamily(function(){ bringAlong(function(moved){
                var left = CUR.data.loss.headless ? CUR.data.loss.leavesFamily.filter(function(x){ return moved.indexOf(x.pid) < 0; }) : [];
                if (left.length) { var d = logLine('bad', 'Still to do in TouchPoint: ' + names(left) + (left.length == 1 ? ' is' : ' are') + ' in a family with no head of household now. Open and add to the right family (Family > Add): ');
                  left.forEach(function(x, i){ d.appendChild(el('a', {href: '/Person2/' + x.pid, target: '_blank', rel: 'noopener', text: (i ? ', ' : '') + x.name})); }); }
                // what to compare next: pairs the deleted record was in (now the record kept), the record
                // kept's other pairs, and the others who look like they are in both families
                var nexts = [], seenN = {};
                function addNext(a, b, text){ var key = Math.min(a, b) + '-' + Math.max(a, b); if (a == b || seenN[key]) return; seenN[key] = 1; nexts.push({a: a, b: b, text: text}); }
                PAIRS.forEach(function(p){ var a = p.keep.pid, b = p.lose.pid;
                  if ((a == k.pid && b == l.pid) || (a == l.pid && b == k.pid)) return;
                  if (a == l.pid || b == l.pid) { var o = a == l.pid ? p.lose : p.keep; addNext(k.pid, o.pid, k.name + ' #' + k.pid + ' and ' + o.name + ' #' + o.pid + ' (it was on the list with the record just deleted)'); }
                  else if (a == k.pid || b == k.pid) addNext(a, b, p.keep.name + ' #' + a + ' and ' + p.lose.name + ' #' + b); });
                var fm = {}; (CUR.data.keepFamily || []).concat(CUR.data.loseFamily || []).forEach(function(p){ fm[p.pid] = p; });
                (CUR.data.familyMatches || []).forEach(function(mm){ var a = fm[mm[0]], b = fm[mm[1]]; if (a && b) addNext(a.pid, b.pid, a.name + ' #' + a.pid + ' and ' + b.name + ' #' + b.pid + ' (in both families)'); });
                if (nexts.length) { var nb = el('div', {'class': 'stack', style: 'gap:6px;margin-top:10px'}, [el('b', {text: 'Compare next'})]);
                  nexts.forEach(function(n){ nb.appendChild(el('div', {'class': 'row'}, [el('span', {style: 'flex:1;min-width:220px', text: n.text}),
                    el('button', {type: 'button', 'class': 'btn ghost', text: 'Compare', onclick: function(){ review(n.a, n.b); }})])); });
                  $('mgdlog').appendChild(nb); }
                PAIRS = PAIRS.filter(function(p){ return p.lose.pid != l.pid && p.keep.pid != l.pid; });
                finish(v.gone && !short.length ? 'Merged, and checked' : 'Merged, with something to check');
              }); });
            });
            });
          });
        }
      });
    });
  }
  // the person kept joins the family the other record was in (TouchPoint's Person > Family > Add)
  function joinFamily(done){
    var k = CUR.data.keep, fam = CUR.data.loss.leavesFamily;
    if (!CUR.joinFamily || !fam.length) { done(); return; }
    CUR.bring = [];
    var m = logLine('busy', 'Moving ' + k.name + ' (#' + k.pid + ') into the family with ' + names(fam) + '...'); m.insertBefore(el('span', {'class': 'spin'}), m.firstChild);
    var x = new XMLHttpRequest(); x.open('POST', '/SearchAdd2/CommitAdd', true);
    x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    x.onload = function(){
      var r = null; try { r = JSON.parse(x.responseText); } catch(e){}
      if (x.status == 200 && r && !r.error) { m.className = 'ok'; m.textContent = k.name + ' (#' + k.pid + ') is now in the family with ' + names(fam) + '.';
        addTo(CUR.bringNew || [], fam[0].pid, 'the family with ' + names(fam), function(){ done(); }); return; }
      else { m.className = 'bad'; m.textContent = 'TouchPoint did not move #' + k.pid + ' into the family' + (r && r.error ? ': ' + r.error : ' (status ' + x.status + ')') + '. Do it in TouchPoint: Person > Family > Add.'; }
      done();
    };
    x.onerror = function(){ m.className = 'bad'; m.textContent = 'Could not reach TouchPoint to move #' + k.pid + ' into the family.'; done(); };
    x.send(['AddContext=family', 'PrimaryKeyForContextType=' + fam[0].pid, encodeURIComponent('PendingList[0].PeopleId') + '=' + k.pid].join('&'));
  }
  // TouchPoint's own Person > Family > Add, for a list of people into the family of one person
  function addTo(list, pid, label, done){
    if (!list.length) { done(); return; }
    var m = logLine('busy', 'Adding ' + names(list) + ' to ' + label + '...'); m.insertBefore(el('span', {'class': 'spin'}), m.firstChild);
    var x = new XMLHttpRequest(); x.open('POST', '/SearchAdd2/CommitAdd', true);
    x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    x.onload = function(){ var r = null; try { r = JSON.parse(x.responseText); } catch(e){}
      if (x.status == 200 && r && !r.error) { m.className = 'ok'; m.textContent = names(list) + (list.length == 1 ? ' is' : ' are') + ' now in ' + label + '.'; }
      else { m.className = 'bad'; m.textContent = 'TouchPoint did not add ' + names(list) + (r && r.error ? ': ' + r.error : ' (status ' + x.status + ')') + '. Do it in TouchPoint: Person > Family > Add.'; }
      done(); };
    x.onerror = function(){ m.className = 'bad'; m.textContent = 'Could not reach TouchPoint to add ' + names(list) + '.'; done(); };
    x.send(['AddContext=family', 'PrimaryKeyForContextType=' + pid].concat(list.map(function(y, i){ return encodeURIComponent('PendingList[' + i + '].PeopleId') + '=' + y.pid; })).join('&'));
  }
  // TouchPoint's own Person > Family > Add: the people checked join the kept person's family
  function bringAlong(done){
    var list = (CUR.bring || []).slice(), k = CUR.data.keep;
    if (!list.length) { done([]); return; }
    var m = logLine('busy', 'Adding ' + names(list) + ' to the family of #' + k.pid + '...'); m.insertBefore(el('span', {'class': 'spin'}), m.firstChild);
    var f = ['AddContext=family', 'PrimaryKeyForContextType=' + k.pid].concat(list.map(function(x, i){ return encodeURIComponent('PendingList[' + i + '].PeopleId') + '=' + x.pid; }));
    var x = new XMLHttpRequest(); x.open('POST', '/SearchAdd2/CommitAdd', true);
    x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    x.onload = function(){
      var r = null; try { r = JSON.parse(x.responseText); } catch(e){}
      if (x.status == 200 && r && !r.error) { m.className = 'ok'; m.textContent = names(list) + (list.length == 1 ? ' is' : ' are') + ' now in the family of #' + k.pid + '.'; done(list.map(function(y){ return y.pid; })); }
      else { m.className = 'bad'; m.textContent = 'TouchPoint did not add ' + names(list) + ' to the family' + (r && r.error ? ': ' + r.error : ' (status ' + x.status + ')') + '.'; done([]); }
    };
    x.onerror = function(){ m.className = 'bad'; m.textContent = 'Could not reach TouchPoint to add ' + names(list) + ' to the family.'; done([]); };
    x.send(f.join('&'));
  }
  $('mgdnext').onclick = function(){ show('mglist'); draw(); };

  // ---- the merge log (each merge and Not the same made here, newest first)
  var SNAP = [['birthday', 'Birthday'], ['gender', 'Gender'], ['marital', 'Marital status'], ['member', 'Member status'], ['email', 'Email'], ['email2', 'Email 2'],
    ['cell', 'Cell'], ['work', 'Work phone'], ['home', 'Home phone'], ['address', 'Address'], ['added', 'Added'], ['how', 'How added'], ['joined', 'Joined'],
    ['baptized', 'Baptized'], ['former', 'Former name'], ['family', 'Family']];
  function drawMerges(r){ var box = $('mglog'); box.innerHTML = '';
    if (!r.success) { box.appendChild(el('span', {'class': 'bad', text: r.message})); return; }
    if (!r.on) { box.appendChild(el('div', {'class': 'muted', text: 'The merge log is off until an Admin opens this page: it first makes the log visible only to Admin and Manager2 (Standard extra values).' + (UPD.logErr ? ' Last try: ' + UPD.logErr : '')})); return; }
    box.appendChild(el('div', {'class': 'muted', text: 'Each merge and Not the same made here, newest first. Kept on the person, seen only by Admin and Manager2. It never includes giving.'}));
    if (!r.entries.length) { box.appendChild(el('div', {'class': 'muted', text: 'Nothing yet.'})); return; }
    r.entries.forEach(function(e){
      var head = e.kind == 'notsame' ? 'Marked ' + e.at.name + ' #' + e.at.pid + ' and ' + (e.other || {}).name + ' #' + (e.other || {}).pid + ' not the same'
        : ((e.from || {}).name || 'A record') + ' #' + (e.from || {}).pid + ' merged into ' + e.at.name + ' #' + e.at.pid;
      var det = el('div', {'class': 'detail'}); det.hidden = true;
      var btn = el('button', {type: 'button', 'class': 'linkbtn', text: head}); btn.onclick = function(){ det.hidden = !det.hidden; };
      var st = e.status == 'check' ? el('span', {'class': 'tag bad', style: 'margin-left:6px', text: 'something to check'}) : (e.status == 'started' ? el('span', {'class': 'tag warn', style: 'margin-left:6px', text: 'did not finish'}) : null);
      if (e.kind == 'merge') {
        var f = e.from || {};
        det.appendChild(el('b', {text: 'The deleted record, #' + f.pid + ' ' + (f.name || '')}));
        SNAP.forEach(function(x){ var v = f[x[0]]; if (v && v.length) det.appendChild(el('div', {'class': 'item'}, [el('span', {'class': 'muted', text: x[1] + ': '}), Array.isArray(v) ? v.join('; ') : String(v)])); });
        if ((e.picks || []).length) { det.appendChild(el('b', {style: 'display:block;margin-top:8px', text: 'Choices'}));
          e.picks.forEach(function(x){ det.appendChild(el('div', {'class': 'item', text: x.field + ': ' + (x.value || '(blank)') + ' (' + x.chose + ')'})); }); }
        if (e.move) det.appendChild(el('div', {'class': 'item', text: 'Family: ' + e.move}));
        if ((e.problems || []).length) { det.appendChild(el('b', {style: 'display:block;margin-top:8px', text: 'To check'}));
          e.problems.forEach(function(x){ det.appendChild(el('div', {'class': 'item bad', text: x})); }); }
        if ((e.steps || []).length) { det.appendChild(el('b', {style: 'display:block;margin-top:8px', text: 'What happened'}));
          e.steps.forEach(function(x){ det.appendChild(el('div', {'class': 'item', text: x})); }); }
      }
      box.appendChild(el('div', {}, [el('div', {'class': 'row', style: 'gap:6px'}, [el('span', {'class': 'muted', text: e.t + (e.by ? ' \u00b7 ' + e.by : '')}), btn, st]), det]));
    });
  }
  $('mglogwrap').addEventListener('toggle', function(){ if (!$('mglogwrap').open) return;
    $('mglog').innerHTML = busy('Reading the merge log...'); post({action: 'merges'}, drawMerges); });

  // ---- updates: an Admin is told when a newer TPxi Duplicates is out
  var UPD = __UPD__;
  function newer(pub, mine){ var a = String(pub || '').split('.'), b = String(mine || '').split('.');
    for (var i = 0; i < Math.max(a.length, b.length); i++){ var x = parseInt(a[i], 10) || 0, y = parseInt(b[i], 10) || 0; if (x !== y) return x > y; }
    return false; }
  function scriptName(){ var m = location.pathname.match(/\/PyScript(?:Form)?\/([^\/?#]+)/); return m ? decodeURIComponent(m[1]) : UPD.id; }
  if (UPD.admin) {
    var skip = ''; try { skip = sessionStorage.getItem('mgupdate.later') || ''; } catch(e) {}
    var vx = new XMLHttpRequest(); vx.open('GET', UPD.api + '/script-versions', true); vx.timeout = 5000;
    vx.onload = function(){ try {
      var latest = JSON.parse(vx.responseText)[UPD.id];
      if (!latest || !newer(latest, UPD.version) || skip == latest) return;
      var b = $('mgupdate'); b.innerHTML = '';
      var go = el('button', {type: 'button', 'class': 'btn', text: 'Update now'}), later = el('button', {type: 'button', 'class': 'btn ghost', text: 'Later'}), st = el('span', {'class': 'muted'});
      b.appendChild(el('b', {text: 'A newer TPxi Duplicates is out. '}));
      b.appendChild(document.createTextNode('You have ' + UPD.version + ', the newest is ' + latest + '. Saved marks (Not the same) stay as they are. '));
      b.appendChild(go); b.appendChild(later); b.appendChild(st); b.hidden = false;
      later.onclick = function(){ try { sessionStorage.setItem('mgupdate.later', latest); } catch(e) {} b.hidden = true; };
      go.onclick = function(){
        if (!go.dataset.sure) { go.dataset.sure = '1'; go.textContent = 'Tap again to update to ' + latest; return; }
        go.disabled = true; st.textContent = ' Updating...';
        post({action: 'apply_update', script_name: scriptName()}, function(r){ st.textContent = ' ' + (r.message || '');
          if (r.success) setTimeout(function(){ location.reload(); }, 800); else go.disabled = false; });
      };
    } catch(e) {} };
    vx.send();
  }

  load();
})();
</script>"""


PROTECT_ERROR = ''


def page():
    return REDIRECT_JS + CSS + (HTML % VERSION) + JS.replace('__UPD__', safe_json(
        {'version': VERSION, 'id': DC_SCRIPT_ID, 'api': DC_API_BASE, 'admin': _in_role('Admin'), 'logErr': PROTECT_ERROR}))


# ---------------------------------------------------------------- run
if model.HttpMethod == 'post':
    handle_post()
elif not can_use():
    model.Form = REDIRECT_JS + CSS + '<div id="mg"><h1>TPxi Duplicates</h1><p>You need the Admin or Manager2 role (the roles TouchPoint\'s own merge screen lets delete a duplicate).</p></div>'
else:
    try:
        model.Header = ''
        if _in_role('Admin') and not log_protected():
            try:
                protect_now()
            except Exception as e:
                PROTECT_ERROR = unicode(e)[:200]
        model.Form = page()
    except Exception as e:
        model.Form = REDIRECT_JS + CSS + '<div id="mg"><h1>TPxi Duplicates</h1><p class="bad">Something went wrong: %s</p></div>' % esc(repr(e))
