#roles=Admin
# -*- coding: utf-8 -*-
#----------------------------------------------------------------------
# TPxi Access Audit
#
# Who can reach what, in one place. Answers three questions off live
# queries: everything one person can reach (the termination list), who
# holds a role and what it unlocks, and every live access token.
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
"""
TPxi_AccessAudit - access audit and termination worklist, inside TouchPoint.

Run from /PyScriptForm/TPxi_AccessAudit

Three questions off one set of live queries:
  By person   everything one human can reach. This is the termination list.
  By role     who holds a role, and what it unlocks
  Tokens      every live access token and who owns it

WHAT IT CHANGES
  Nothing, until you select rows, pick an action and confirm it. On one
  person's open task notes it can reassign, complete, archive or delete.
  Reassigning and completing each send an email per note. Archiving and
  deleting send none, and only archiving can be undone. It can also remove
  a role, but only from somebody with a single login, and revoke access
  tokens. Every action is scoped to what you selected.

TWO THINGS DELIBERATELY NOT USED
  model.TaskNoteMassComplete sets taskNote.Notes = note on every row, which
  overwrites the existing note. 138,499 open task notes here have real content
  in that field, so a bulk close through it would destroy pastoral history.
  model.TaskNoteComplete has the same flaw, taskNote.Notes = note with no null
  check, so an empty note blanks it. This reads the existing note first and
  hands the same text back, which makes that assignment a no-op and needs no
  token. Reassigning needs no token either. The only thing that does is
  revoking access tokens, because TouchPoint exposes no model API for it.

  The role to code map cannot be built in here, it needs the TouchPoint C#
  source. It is generated locally and stored as Special Content. If it is
  missing the By role tab still works, it just shows holders without the
  "what it unlocks" detail.

Upload as Special Content > Python, name TPxi_AccessAudit.
Needs Admin. Add to CustomReports with role="Admin".
"""

import datetime
import re
import traceback

APP_VERSION = "1.0.3"
ROLE_MAP_CONTENT = "TPxi_AccessAudit_RoleMap"

# --- Auto update, see TPxi/AutoUpdate/README.md ----------------------------
# scripts.displaycache.com is the browser facing version check. The workers.dev
# mirror is what the SERVER fetches from, because Cloudflare Bot Fight Mode lets
# browsers through and blocks server side requests.
DC_SCRIPT_ID = "TPxi_AccessAudit"
DC_API_BASE = "https://scripts.displaycache.com/api/touchpoint"
DC_API_WORKER = "https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint"
# TaskNote.StatusId: 1 Complete, 2 Pending, 3 Accepted, 4 Declined, 5 Note,
# 6 Archived. Only Pending and Accepted are live work. Filtering on "not
# Complete" looks right and is badly wrong: 136,953 of the 141,588 rows that
# passed it are Archived. It made one person look like they had left 3,701
# open tasks behind when the real number was zero.
# TouchPoint only writes LastLoginDate on some sign in paths. 9,766 of the
# 11,612 user rows here have none at all, yet 7,224 of those do have a
# LastActivityDate, and 1,775 more have activity newer than their last recorded
# login. Judging staleness on LastLoginDate alone called 332 role holders
# dormant who had been active that same year. Always take the later of the two.
SEEN = ("CASE WHEN ISNULL(u.LastActivityDate, '1900-01-01') > "
        "ISNULL(u.LastLoginDate, '1900-01-01') "
        "THEN u.LastActivityDate ELSE u.LastLoginDate END")

OPEN_TASK = "tn.IsNote = 0 AND tn.StatusId IN (2,3) AND ISNULL(tn.IsArchived,0) = 0"
MAX_ROWS = 400          # per list, keeps the page honest on a big church

# ---------------------------------------------------------------------------
# json, written by hand. IronPython's json.dumps dies on the non-ASCII that
# arrives constantly in names and involvement titles
# ---------------------------------------------------------------------------
def _esc(s):
    if not isinstance(s, (str, unicode)):
        s = unicode(s)
    out = ['"']
    for ch in s:
        try:
            code = ord(ch)
        except Exception:
            out.append('\\ufffd'); continue
        if   code == 0x22: out.append('\\"')
        elif code == 0x5C: out.append('\\\\')
        elif code == 0x0A: out.append('\\n')
        elif code == 0x0D: out.append('\\r')
        elif code == 0x09: out.append('\\t')
        elif code < 0x20 or code >= 0x7F: out.append('\\u%04x' % code)
        else: out.append(chr(code))
    out.append('"')
    return ''.join(out)

def _enc(o):
    if o is None: return 'null'
    if o is True: return 'true'
    if o is False: return 'false'
    if isinstance(o, (int, long)): return str(o)
    if isinstance(o, float): return 'null' if o != o else repr(o)
    if isinstance(o, (str, unicode)): return _esc(o)
    if isinstance(o, datetime.datetime): return _esc(o.isoformat())
    if isinstance(o, datetime.date): return _esc(o.isoformat())
    if isinstance(o, dict):
        return '{' + ','.join(_esc(unicode(k)) + ':' + _enc(v) for k, v in o.items()) + '}'
    if isinstance(o, (list, tuple)):
        return '[' + ','.join(_enc(x) for x in o) + ']'
    return _esc(unicode(o))

def safe_json(o):
    try:
        return _enc(o)
    except Exception as e:
        return '{"ok":false,"message":' + _esc('safe_json: ' + repr(e)) + '}'

# ---------------------------------------------------------------------------
def sval(row, name, default=None):
    try:
        v = getattr(row, name)
        return default if v is None else v
    except Exception:
        return default

def _parts(v):
    """Year, month, day out of whatever the row handed us.

    This is the thing that quietly broke every date in this report. A .NET
    System.DateTime exposes Year, Month and Day with capitals. Lowercase
    v.year raises AttributeError, which a bare except turns into None, so
    every date rendered as "no activity recorded" including for people who
    had signed in that morning. Nothing reproduces it outside IronPython,
    because a real python datetime answers to the lowercase names.

    Order matters: try the .NET names first, then python's, then fall back to
    parsing the string, which TouchPoint formats US style as M/D/YYYY."""
    for hi, lo in (("Year", "Month"), ("year", "month")):
        try:
            y = getattr(v, hi)
            mo = getattr(v, lo)
            d = getattr(v, "Day" if hi == "Year" else "day")
            return int(y), int(mo), int(d)
        except Exception:
            pass
    try:
        head = str(v).strip().split(" ")[0]
        if "/" in head:                      # 9/15/2026
            mo, d, y = head.split("/")[:3]
            return int(y), int(mo), int(d)
        if "-" in head:                      # 2026-09-15
            y, mo, d = head.split("-")[:3]
            return int(y), int(mo), int(d)
    except Exception:
        pass
    return None


def datestr(v):
    """ISO date, or None. Never slice str(v): TouchPoint formats dates US style
    so str(dt)[:10] eats into the time on single digit months and days."""
    if v is None:
        return None
    p = _parts(v)
    return None if p is None else "%04d-%02d-%02d" % p

def today():
    return datetime.datetime.now()

def days_since(d):
    if not d:
        return None
    p = _parts(d)
    if p is None:
        return None
    try:
        return (today() - datetime.datetime(p[0], p[1], p[2])).days
    except Exception:
        return None

# ---------------------------------------------------------------------------
# Auto update. See TPxi/AutoUpdate/README.md.
# ---------------------------------------------------------------------------
def get_script_name():
    """What this script is actually installed as, which may not be DC_SCRIPT_ID
    if an admin renamed it. Posted name first, then the URL, then the default."""
    try:
        v = getattr(model.Data, "script_name", None)
        if v and str(v).strip():
            return str(v).strip()
    except Exception:
        pass
    try:
        url = str(getattr(model, "URL", "") or "")
        m = re.search(r"/PyScript(?:Form)?/([^/?#&]+)", url)
        if m:
            return m.group(1)
    except Exception:
        pass
    return DC_SCRIPT_ID


def do_apply_update():
    """Pull the published source and overwrite this script's content slot.

    Nothing this tool shows is kept in the script. Roles, logins and tokens are
    TouchPoint's own tables, the role map is its own content record, and the
    one setting lives in Admin. So replacing the code loses nothing."""
    try:
        code = str(model.RestGet(DC_API_WORKER + "/scripts/" + DC_SCRIPT_ID, {}))
    except Exception as e:
        return {"ok": False, "message": "Could not fetch the update: " + str(e)[:200]}
    if not code or len(code) < 200:
        return {"ok": False,
                "message": "The update came back empty or too short to be real"}
    if "APP_VERSION" not in code:
        return {"ok": False,
                "message": "That did not look like this script, so nothing was written"}
    target = get_script_name() or DC_SCRIPT_ID
    try:
        model.WriteContentPython(target, code)
    except Exception as e:
        return {"ok": False, "message": "Write failed: " + str(e)[:200]}
    return {"ok": True, "message": "Updated " + target + ". Reloading."}


# ---------------------------------------------------------------------------
# Who is running this. UserPeopleId is a property; calling it raises, and a
# bare except then gives every user the same identity.
# ---------------------------------------------------------------------------
def current_user_id():
    try:
        uid = model.UserPeopleId
    except Exception:
        uid = None
    if not uid or int(uid) <= 0:
        return None
    return int(uid)


# ===========================================================================
# Queries.
#
# Everything here is set based on purpose. The obvious shape, one SELECT over
# People with a correlated subquery per person, never returned on this data:
# it runs the TaskNote scan once per person across 11k people. Aggregating each
# source once and joining the results in python takes about half a second.
# ===========================================================================

def _chunks(seq, n=800):
    seq = list(seq)
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def last_activity_by_user():
    """UserId -> the newest thing ActivityLog saw them do.

    Users.LastLoginDate and Users.LastActivityDate are not written on every
    auth path. On some installs they are null even for people signing in daily,
    which made this report announce that its own author had never signed in.
    ActivityLog records the actual page hits, so it is the one source that
    cannot disagree with reality. Bounded to 400 days because the table is
    large."""
    out = {}
    try:
        for r in q.QuerySql("""
                SELECT a.UserId, MAX(a.ActivityDate) AS Last
                FROM dbo.ActivityLog a WITH (NOLOCK)
                WHERE a.UserId IS NOT NULL
                  AND a.ActivityDate >= DATEADD(day, -400, GETDATE())
                GROUP BY a.UserId"""):
            uid = sval(r, "UserId")
            if uid:
                out[uid] = sval(r, "Last")
    except Exception:
        pass
    return out


def newer(a, b):
    """Later of two dates, either of which may be missing.

    Compared on extracted parts rather than directly, because one side can be a
    .NET DateTime and the other a python one, and those do not compare."""
    if a is None:
        return b
    if b is None:
        return a
    pa, pb = _parts(a), _parts(b)
    if pa is None:
        return b
    if pb is None:
        return a
    return a if pa >= pb else b


def build_index():
    """One pass over each source. Returns pid -> footprint."""
    idx = {}
    seen_log = last_activity_by_user()

    def slot(pid):
        pid = int(pid)
        if pid not in idx:
            idx[pid] = {"pid": pid, "name": "", "logins": 0, "roles": 0, "tokens": 0,
                        "last": None, "locked": 0, "tasks_own": 0, "tasks_asg": 0,
                        "notify": 0, "giftnotify": 0, "regfrom": 0, "led": 0,
                        "deceased": False, "archived": False, "nobdate": False}
        return idx[pid]

    # logins, and the newest sign in across them
    for r in q.QuerySql("""
            SELECT u.PeopleId, u.UserId,
                   %s AS LastSeen,
                   u.IsLockedOut
            FROM dbo.Users u WITH (NOLOCK)
            WHERE u.PeopleId IS NOT NULL""" % SEEN):
        pid = sval(r, "PeopleId")
        if not pid:
            continue
        s_ = slot(pid)
        s_["logins"] = s_.get("logins", 0) + 1
        if sval(r, "IsLockedOut", False):
            s_["locked"] = s_.get("locked", 0) + 1
        best = newer(sval(r, "LastSeen"), seen_log.get(sval(r, "UserId")))
        s_["lastraw"] = newer(s_.get("lastraw"), best)
        s_["last"] = datestr(s_["lastraw"])

    for r in q.QuerySql("""
            SELECT u.PeopleId, COUNT(*) AS N
            FROM dbo.UserRole ur WITH (NOLOCK)
            JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = ur.UserId
            WHERE u.PeopleId IS NOT NULL
            GROUP BY u.PeopleId"""):
        slot(sval(r, "PeopleId"))["roles"] = sval(r, "N", 0)

    for r in q.QuerySql("""
            SELECT u.PeopleId, COUNT(*) AS N
            FROM dbo.UserTokens t WITH (NOLOCK)
            JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = t.UserId
            WHERE u.PeopleId IS NOT NULL
            GROUP BY u.PeopleId"""):
        slot(sval(r, "PeopleId"))["tokens"] = sval(r, "N", 0)

    for fld, key in (("OwnerId", "tasks_own"), ("AssigneeId", "tasks_asg")):
        for r in q.QuerySql("""
                SELECT {0} AS Pid, COUNT(*) AS N
                FROM dbo.TaskNote tn WITH (NOLOCK)
                WHERE {1} AND {0} IS NOT NULL
                GROUP BY {0}""".format(fld, OPEN_TASK)):
            pid = sval(r, "Pid")
            if pid:
                slot(pid)[key] = sval(r, "N", 0)

    # one pass over involvements; NotifyIds is a comma list so it is split here
    orgs = []
    for r in q.QuerySql("""
            SELECT o.OrganizationId, o.OrganizationName, o.OrganizationStatusId,
                   ISNULL(o.RegistrationTypeId,0) AS RegType,
                   ISNULL(o.RegMessageNotificationEmailSend,0) AS NSend,
                   ISNULL(o.RegMessageConfirmationEmailSend,0) AS CSend,
                   o.LeaderId, o.RegMessageConfirmationEmailFrom,
                   ISNULL(o.NotifyIds,'') AS NotifyIds,
                   ISNULL(o.GiftNotifyIds,'') AS GiftNotifyIds
            FROM dbo.Organizations o WITH (NOLOCK)"""):
        o = {"oid": sval(r, "OrganizationId", 0),
             "name": sval(r, "OrganizationName", "") or "",
             "inactive": sval(r, "OrganizationStatusId", 30) != 30,
             "regtype": sval(r, "RegType", 0),
             "nsend": bool(sval(r, "NSend", 0)),
             "csend": bool(sval(r, "CSend", 0)),
             "leader": sval(r, "LeaderId"),
             "from": sval(r, "RegMessageConfirmationEmailFrom"),
             "notify": [x for x in str(sval(r, "NotifyIds", "")).replace(" ", "").split(",") if x.isdigit()],
             "gift": [x for x in str(sval(r, "GiftNotifyIds", "")).replace(" ", "").split(",") if x.isdigit()]}
        orgs.append(o)
        if o["leader"]:
            slot(o["leader"])["led"] += 1
        if o["from"]:
            slot(o["from"])["regfrom"] += 1
        for x in o["notify"]:
            slot(int(x))["notify"] += 1
        for x in o["gift"]:
            slot(int(x))["giftnotify"] += 1

    # names and status, only for the people who turned up
    for batch in _chunks(idx.keys()):
        ids = ",".join(str(x) for x in batch)
        for r in q.QuerySql("""
                SELECT p.PeopleId, p.Name2, p.IsDeceased, p.ArchivedFlag,
                       CASE WHEN p.BDate IS NULL THEN 1 ELSE 0 END AS NoBDate,
                       CASE WHEN ISNULL(p.EmailAddress,'') = ''
                             AND ISNULL(p.EmailAddress2,'') = '' THEN 1 ELSE 0 END AS NoEmail
                FROM dbo.People p WITH (NOLOCK)
                WHERE p.PeopleId IN (%s)""" % ids):
            pid = sval(r, "PeopleId")
            if pid in idx:
                s_ = idx[pid]
                s_["name"] = sval(r, "Name2", "") or ("PeopleId %s" % pid)
                s_["deceased"] = bool(sval(r, "IsDeceased", False))
                s_["archived"] = bool(sval(r, "ArchivedFlag", False))
                s_["nobdate"] = bool(sval(r, "NoBDate", 0))
                s_["noemail"] = bool(sval(r, "NoEmail", 0))

    # a login with nothing attached is not a footprint
    keep = {}
    for pid, v in idx.items():
        if (v["roles"] or v["tokens"] or v["tasks_own"] or v["tasks_asg"]
                or v["notify"] or v["giftnotify"] or v["regfrom"] or v["led"]):
            v["footprint"] = (v["notify"] + v["giftnotify"] + v["regfrom"] + v["led"])
            v["idledays"] = days_since(v.get("lastraw"))
            # no birthdate usually means a department address or a service
            # account rather than a person. A hint, not proof: one real person
            # here has none, so it only demotes and labels, never excludes.
            v["service"] = bool(v["nobdate"])
            v.pop("lastraw", None)
            if not v["name"]:
                v["name"] = "PeopleId %s" % pid
            keep[pid] = v
    return keep, orgs


def person_orgs(pid, orgs):
    """Slice the single involvement pass down to one person."""
    pid = int(pid)
    out = {"notify": [], "giftnotify": [], "regfrom": [], "led": []}
    sp = str(pid)
    for o in orgs:
        row = {"oid": o["oid"], "name": o["name"], "inactive": o["inactive"],
               "state": "" if (o["regtype"] and o["nsend"]) else
                        ("no registration" if not o["regtype"] else "email off")}
        if sp in o["notify"]:
            out["notify"].append(row)
        if sp in o["gift"]:
            g = dict(row); g["state"] = ""
            out["giftnotify"].append(g)
        if o["from"] == pid:
            f = dict(row)
            f["state"] = "" if (o["regtype"] and o["csend"]) else \
                         ("no registration" if not o["regtype"] else "email off")
            out["regfrom"].append(f)
        if o["leader"] == pid:
            l = dict(row); l["state"] = ""
            out["led"].append(l)
    return out


def person_logins(pid):
    pid = int(pid)
    logins = {}
    seen_log = last_activity_by_user()
    for r in q.QuerySql("""
            SELECT u.UserId, u.Username, {0} AS LastSeen, u.IsLockedOut,
                   (SELECT COUNT(*) FROM dbo.UserTokens t WITH (NOLOCK)
                    WHERE t.UserId = u.UserId) AS Toks
            FROM dbo.Users u WITH (NOLOCK)
            WHERE u.PeopleId = {1} ORDER BY u.UserId""".format(SEEN, pid)):
        uid = sval(r, "UserId", 0)
        logins[uid] = {"uid": uid, "user": sval(r, "Username", "") or "",
                       "last": datestr(newer(sval(r, "LastSeen"), seen_log.get(uid))),
                       "locked": bool(sval(r, "IsLockedOut", False)),
                       "tokens": sval(r, "Toks", 0), "roles": []}
    if logins:
        ids = ",".join(str(x) for x in logins.keys())
        for r in q.QuerySql("""
                SELECT ur.UserId, r.RoleName
                FROM dbo.UserRole ur WITH (NOLOCK)
                JOIN dbo.Roles r WITH (NOLOCK) ON r.RoleId = ur.RoleId
                WHERE ur.UserId IN (%s) ORDER BY r.RoleName""" % ids):
            u = logins.get(sval(r, "UserId", 0))
            if u:
                u["roles"].append(sval(r, "RoleName", ""))
    return [logins[k] for k in sorted(logins.keys())]


def person_leadmem(pid):
    out = []
    for r in q.QuerySql("""
            SELECT TOP 200 o.OrganizationId, o.OrganizationName, o.OrganizationStatusId,
                   mt.Description AS MT
            FROM dbo.OrganizationMembers om WITH (NOLOCK)
            JOIN dbo.Organizations o WITH (NOLOCK) ON o.OrganizationId = om.OrganizationId
            LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
            WHERE om.PeopleId = %d AND om.MemberTypeId IN (140,310,320)
            ORDER BY o.OrganizationName""" % int(pid)):
        out.append({"oid": sval(r, "OrganizationId", 0),
                    "name": sval(r, "OrganizationName", "") or "",
                    "inactive": sval(r, "OrganizationStatusId", 30) != 30,
                    "mt": sval(r, "MT", "Leader")})
    return out


# ---------------------------------------------------------------------------
# What they have actually been doing. Two separate trails, kept separate
# because they answer different questions and only one of them is complete.
# ---------------------------------------------------------------------------

# Activity text is free form, so this classifies on what the strings really
# look like in ActivityLog rather than on what they ought to look like. Every
# pattern below was checked against real rows before it went in. Order counts:
# the first match wins, so the narrow ones come before the broad ones.
ACT_KINDS = [
    ("role write", "Someone's roles were changed from a script",
     ("%AddRole%", "%RemoveRole%", "%SetRoles%")),
    ("delete", "A person, involvement, meeting or task was deleted",
     ("Delete%", "%deleted by%", "Deleted %")),
    ("merge", "Two records were merged, which is not reversible",
     ("Merg%",)),
    ("giving", "Contribution or statement data was opened",
     ("%Contribution%", "%eStatement%", "%Giving%", "%Bundle%")),
    ("download", "A file was pulled out of TouchPoint, including bank deposit files",
     ("Downloading%", "%Export%", "%Excel%")),
    ("bulk email", "Email was sent to a group rather than one person",
     ("Emailing%",)),
    ("script", "A Python script was run",
     ("Run Python script%",)),
    ("auth failure", "A sign in was refused",
     ("Invalid log in%", "%ailed password%", "ForgotPassword%")),
]


def _act_case(col):
    """The CASE that buckets an activity string. Built once here so the summary
    and the detail cannot drift apart and disagree about what counts."""
    out = []
    for kind, _desc, pats in ACT_KINDS:
        cond = " OR ".join("%s LIKE '%s'" % (col, p) for p in pats)
        out.append("WHEN %s THEN '%s'" % (cond, kind))
    return "CASE " + " ".join(out) + " ELSE 'routine' END"


def person_changes(pid, days):
    """Records this person edited, newest first.

    ChangeLog is the header, one row per save, and ChangeDetails holds the
    field level before and after. TouchPoint's own note that ChangeDetails is
    unused is wrong: it has 658,262 rows here and is the only place the actual
    field names live.

    UserPeopleId is who made the change. It is 0 on roughly half of recent
    rows, which is a self service edit or a kiosk rather than a staff member,
    and those are simply not attributable to anybody. So this never reports
    zero as a person, and an empty result means "nothing attributed", not
    "they changed nothing"."""
    pid = int(pid)
    days = int(days)
    if pid <= 0:
        return []
    out = []
    try:
        rows = q.QuerySql("""
            WITH ev AS (
                SELECT TOP 300 cl.Id, cl.Created, cl.PeopleId AS OnId, cl.Field AS Area
                FROM dbo.ChangeLog cl WITH (NOLOCK)
                WHERE cl.UserPeopleId = %d
                  AND cl.Created >= DATEADD(day, -%d, GETDATE())
                ORDER BY cl.Id DESC
            ), det AS (
                SELECT cd.Id,
                       STRING_AGG(CAST(cd.Field AS NVARCHAR(MAX)), ', ') AS Flds,
                       COUNT(*) AS NF
                FROM dbo.ChangeDetails cd WITH (NOLOCK)
                WHERE cd.Id IN (SELECT Id FROM ev)
                GROUP BY cd.Id
            )
            SELECT ev.Id, ev.Created, ev.OnId, ev.Area,
                   ISNULL(p.Name, '') AS OnName,
                   ISNULL(det.Flds, '') AS Flds, ISNULL(det.NF, 0) AS NF
            FROM ev
            LEFT JOIN dbo.People p WITH (NOLOCK) ON p.PeopleId = ev.OnId
            LEFT JOIN det ON det.Id = ev.Id
            ORDER BY ev.Id DESC""" % (pid, days))
        for r in rows:
            out.append({
                "id": sval(r, "Id"),
                "when": datestr(sval(r, "Created")),
                "onid": sval(r, "OnId"),
                "on": sval(r, "OnName", "") or ("PeopleId %s" % sval(r, "OnId")),
                "area": sval(r, "Area", ""),
                "flds": sval(r, "Flds", ""),
                "nf": sval(r, "NF", 0),
                "self": sval(r, "OnId") == pid,
            })
    except Exception:
        pass
    return out


def person_activity(pid, days):
    """What this person did in TouchPoint, from ActivityLog.

    Joined through Users on PeopleId so it covers every login they own, not
    just the first, which is the mistake this whole report exists to point at.

    Two halves on purpose. The counts cover everything. The list is only the
    kinds worth reading one by one, because a busy staff member generates
    thousands of routine page hits a month and burying six deletes in them
    helps nobody."""
    pid = int(pid)
    days = int(days)
    if pid <= 0:
        return {"counts": [], "rows": [], "total": 0}
    case = _act_case("al.Activity")
    counts, rows, total = [], [], 0
    try:
        for r in q.QuerySql("""
                SELECT %s AS Kind, COUNT(*) AS N
                FROM dbo.ActivityLog al WITH (NOLOCK)
                JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = al.UserId
                WHERE u.PeopleId = %d
                  AND al.ActivityDate >= DATEADD(day, -%d, GETDATE())
                GROUP BY %s""" % (case, pid, days, case)):
            n = sval(r, "N", 0) or 0
            total += n
            counts.append({"kind": sval(r, "Kind", ""), "n": n})
    except Exception:
        pass
    try:
        for r in q.QuerySql("""
                SELECT TOP 250 al.ActivityDate, al.Activity, al.PeopleId,
                       %s AS Kind
                FROM dbo.ActivityLog al WITH (NOLOCK)
                JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = al.UserId
                WHERE u.PeopleId = %d
                  AND al.ActivityDate >= DATEADD(day, -%d, GETDATE())
                  AND %s <> 'routine'
                ORDER BY al.ActivityDate DESC""" % (case, pid, days, case)):
            rows.append({
                "when": datestr(sval(r, "ActivityDate")),
                "what": sval(r, "Activity", ""),
                "onid": sval(r, "PeopleId"),
                "kind": sval(r, "Kind", ""),
            })
    except Exception:
        pass
    counts.sort(key=lambda c: -c["n"])
    return {"counts": counts, "rows": rows, "total": total,
            "kinds": [{"k": k, "d": d} for k, d, _ in ACT_KINDS]}


def open_task_rows(pid, cap=200):
    """The tasks themselves, so you can see what you are about to move."""
    pid = int(pid)
    sql = """
            SELECT TOP {0} tn.TaskNoteId, tn.StatusId, tn.DueDate, tn.CreatedDate,
                   tn.OwnerId, tn.AssigneeId, tn.AboutPersonId,
                   ab.Name2 AS AboutName,
                   LEFT(ISNULL(tn.Instructions,''), 170) AS Instr,
                   o.OrganizationName AS OrgName,
                   STUFF((SELECT ', ' + k.Description
                          FROM dbo.TaskNoteKeyword tk WITH (NOLOCK)
                          JOIN dbo.Keyword k WITH (NOLOCK) ON k.KeywordId = tk.KeywordId
                          WHERE tk.TaskNoteId = tn.TaskNoteId
                          FOR XML PATH('')), 1, 2, '') AS Keywords
            FROM dbo.TaskNote tn WITH (NOLOCK)
            LEFT JOIN dbo.People ab WITH (NOLOCK) ON ab.PeopleId = tn.AboutPersonId
            LEFT JOIN dbo.Organizations o WITH (NOLOCK) ON o.OrganizationId = tn.OrgId
            WHERE {1} AND (tn.OwnerId = {2} OR tn.AssigneeId = {2})
            ORDER BY ISNULL(tn.DueDate, tn.CreatedDate)""".format(cap, OPEN_TASK, pid)
    out = []
    for r in q.QuerySql(sql):
        role = []
        if sval(r, "OwnerId") == pid:
            role.append("owner")
        if sval(r, "AssigneeId") == pid:
            role.append("assignee")
        out.append({"id": sval(r, "TaskNoteId", 0),
                    "status": "Accepted" if sval(r, "StatusId") == 3 else "Pending",
                    "role": " and ".join(role),
                    "about": sval(r, "AboutName", "") or "",
                    "aboutid": sval(r, "AboutPersonId", 0),
                    "due": datestr(sval(r, "DueDate")),
                    "made": datestr(sval(r, "CreatedDate")),
                    "what": (sval(r, "Instr", "") or "").strip(),
                    "org": sval(r, "OrgName", "") or "",
                    "kw": sval(r, "Keywords", "") or ""})
    return out


def open_task_ids(pid, which, cap=5000):
    """Ids only. which is owner, assignee or both."""
    pid = int(pid)
    if which == "owner":
        cond = "tn.OwnerId = %d" % pid
    elif which == "assignee":
        cond = "tn.AssigneeId = %d" % pid
    else:
        cond = "(tn.OwnerId = %d OR tn.AssigneeId = %d)" % (pid, pid)
    out = []
    for r in q.QuerySql("""
            SELECT TOP {0} tn.TaskNoteId
            FROM dbo.TaskNote tn WITH (NOLOCK)
            WHERE {1} AND {2}
            ORDER BY tn.TaskNoteId""".format(cap, OPEN_TASK, cond)):
        t = sval(r, "TaskNoteId", 0)
        if t:
            out.append(t)
    return out


# ===========================================================================
# Reference data that cannot be computed in here
# ===========================================================================

def role_map():
    """role -> what it unlocks.

    This is the one thing the script cannot work out for itself: "what does the
    Finance role unlock" is a question about TouchPoint's own C# source, not
    about this database. So it is generated by _rolemap/scan_source.py and baked
    in at the bottom of this file, which keeps the whole tool to one upload.

    Special Content named TPxi_AccessAudit_RoleMap overrides it if present, so a
    refreshed map can be dropped in after a TouchPoint release without touching
    the script."""
    import json
    try:
        raw = model.TextContent(ROLE_MAP_CONTENT)
        if raw and raw.strip().startswith("{"):
            return json.loads(raw)
    except Exception:
        pass
    try:
        return json.loads(ROLE_MAP_JSON)
    except Exception:
        return None


def role_screen_settings():
    """Every role's 60 screen settings, read live out of Content.

    TouchPoint keeps these in a Content record called CustomAccessRoles.xml,
    one block per role, every setting written out explicitly. The friendly
    names, the tooltips and the Involvement / Meetings / General / Person
    grouping are not in there: they live in TouchPoint's own C# resource,
    RoleSettingDefaults.xml, which is what ROLE_SETTINGS_JSON at the bottom of
    this file is a copy of. So the wording you see here is TouchPoint's, not
    mine.

    Returns (order, values). Order matters more than it looks: see
    setting_owner() below."""
    order, values = [], {}
    try:
        rows = q.QuerySql("""
            SELECT TOP 1 c.Body FROM dbo.Content c WITH (NOLOCK)
            WHERE c.Name = 'CustomAccessRoles.xml'""")
        body = sval(rows[0], "Body", "") if rows else ""
    except Exception:
        body = ""
    if not body:
        return order, values
    # A regex rather than an XML parser on purpose. The document is machine
    # written by TouchPoint and never hand edited, and IronPython's XML support
    # is not something to bet a 350KB parse on inside a page render.
    try:
        for m in re.finditer(r'<role\s+name="([^"]*)"(.*?)</role>', body, re.S):
            name = m.group(1)
            if not name:
                continue
            order.append(name)
            cur = {}
            for sm in re.finditer(r'<setting\s+name="([^"]*)"\s+value="([^"]*)"',
                                  m.group(2)):
                cur[sm.group(1)] = sm.group(2).strip().lower() == "true"
            values[name] = cur
    except Exception:
        return [], {}
    return order, values


def setting_catalog():
    """The 60 settings with TouchPoint's own labels, tooltips and grouping."""
    import json
    try:
        return json.loads(ROLE_SETTINGS_JSON)
    except Exception:
        return []


def setting_owner(order, values, roles, key):
    """Which of a person's roles actually decides one setting.

    This is the part worth knowing. RoleChecker.HasSetting walks the roles in
    the order they appear in CustomAccessRoles.xml and returns the FIRST one
    the user holds that has a value. It does not take the most permissive, and
    it does not combine them. Every role carries an explicit value for every
    setting, so the first role in that order wins outright and the rest are
    never consulted.

    The order is Roles.Priority ascending, with MyData first. 30 of the 87
    roles here have no priority set at all, and a null sorts ahead of every
    number, so the roles nobody ever ordered outrank the ones that were
    deliberately ordered. That is worth seeing before you read a value off
    this screen and assume it is what somebody gets."""
    have = set(roles or [])
    for rn in order:
        if rn in have and key in values.get(rn, {}):
            return rn
    return None


# ---------------------------------------------------------------------------
# Which of your own scripts mention a role.
#
# This is the thing the "nothing found" message tells you to go and check by
# hand, and it is the only place a custom role usually turns up: TouchPoint's
# code never checks a custom role, so if IT-Team does anything, a script here
# is doing it.
#
# Why it is cached rather than live. The work is O(roles x bytes). Doing it in
# SQL costs a full pass over every content body PER ROLE: one role took 7
# seconds against 64MB here. Pulling the bodies once and testing every role
# against each in memory does the whole corpus in under 5. That is fast enough
# to run, and far too slow to run on every page load, so it runs in batches on
# demand and the answer is kept in Special Content.
#
# Batching is what makes it portable. A church with ten times this content
# would take fifty seconds in one request and time out; in batches it just
# takes more batches.
# ---------------------------------------------------------------------------
SCRIPT_SCAN_CONTENT = "TPxi_AccessAudit_ScriptScan"
SCRIPT_SCAN_BATCH = 120
SCRIPT_SCAN_CAP = 12          # per role, so one role cannot bloat the cache
# CustomAccessRoles.xml is TouchPoint's own per-role settings file: it names
# every role that exists, so scanning it hands all 87 a meaningless hit. The
# cache record is skipped for the same reason once it has been written.
SCRIPT_SCAN_SKIP = ("CustomAccessRoles.xml", SCRIPT_SCAN_CONTENT)


def script_scan_cache():
    """What the last scan found, or None if it has never been run."""
    import json
    try:
        raw = model.TextContent(SCRIPT_SCAN_CONTENT)
        if raw and raw.strip().startswith("{"):
            d = json.loads(raw)
            if isinstance(d.get("roles"), dict):
                return d
    except Exception:
        pass
    return None


SCRIPT_SCAN_MAX_AGE_HOURS = 24


def script_scan_age_hours():
    """Hours since the cache was written, or None if there is no cache.

    Kept on the server so that five admins opening the page at nine o'clock
    cannot each start their own scan. The browser asks; the server decides."""
    import datetime as _dt
    c = script_scan_cache()
    if not c or not c.get("when"):
        return None
    try:
        w = _dt.datetime.strptime(str(c["when"])[:16], "%Y-%m-%d %H:%M")
    except Exception:
        return None
    try:
        d = _dt.datetime.now() - w
        return (d.days * 86400 + d.seconds) / 3600.0
    except Exception:
        return None


def script_scan_needed():
    a = script_scan_age_hours()
    return a is None or a >= SCRIPT_SCAN_MAX_AGE_HOURS


def scan_base(nm):
    """(base name, is this a backup copy).

    Content is where people keep their working copies, so one script shows up
    as a dozen records: TPxi_QuickLinksAdmin, TPxi_QuickLinksAdmin20260326,
    QuickLinks_Bak_20260602_065420. Listing all of them buries the live script
    and, worse, fills the per-role cap so a real script gets dropped.

    Stripped repeatedly because the markers stack: a name can carry a Bak tag
    and two dates. Only suffixes are touched; anything left of them is the
    name somebody actually chose."""
    out, backup, prev = nm, False, None
    while prev != out:
        prev = out
        for rx in (r'[ _-]?\d{4}-\d{2}-\d{2}$',          # _2026-06-01
                   r'[ _-]?\d{8}(_\d{6})?$',              # 20260326, _20260602_065420
                   r'[ _-]?[Bb][Aa][Kk](up)?[ _-]?$',      # _Bak, -backup
                   r'[ _-]?([Oo][Ll][Dd]|[Cc]opy)$'):      # _old, _Copy
            n2 = re.sub(rx, "", out)
            if n2 != out and n2:
                out, backup = n2, True
    # a Bak tag anywhere marks it, eg MenuEditor_Bak_ReportsMenuPeople.xml_...
    if re.search(r'[ _-][Bb][Aa][Kk][ _-]', nm):
        backup = True
    return out, backup


def scan_add(found, rn, nm, kind):
    """Record a hit, folding backup copies into the live script.

    Entry shape is [display name, kind, how many older copies]. The live copy
    wins the display slot whenever one exists; where only backups survive, the
    base name is shown so it still reads as a name rather than a timestamp."""
    base, is_bak = scan_base(nm)
    lst = found.setdefault(rn, [])
    for e in lst:
        if e[0] == base or (len(e) > 3 and e[3] == base):
            if is_bak or e[0] != nm:
                e[2] = e[2] + 1
            if not is_bak:
                e[0] = nm                      # a real copy outranks a backup
            return
    if len(lst) >= SCRIPT_SCAN_CAP:
        return
    lst.append([base if is_bak else nm, kind, 0, base])


def script_scan_total():
    """How many content records there are to scan, or -1 if that cannot be
    determined.

    It returns -1 rather than 0 on failure on purpose. A swallowed error that
    comes back as 0 looked exactly like "nothing to scan": the batch loop saw
    120 >= 0, decided it had finished after the first batch, and saved a
    partial answer that read "120 of 0 content records". The count is only a
    progress figure now, never the thing that decides when to stop."""
    try:
        for r in q.QuerySql("""
                SELECT COUNT(*) AS N FROM dbo.Content c WITH (NOLOCK)
                WHERE c.TypeID IN (1,5) AND c.Name NOT IN (%s)"""
                % ",".join("'" + x.replace("'", "''") + "'" for x in SCRIPT_SCAN_SKIP)):
            return int(sval(r, "N", 0) or 0)
    except Exception:
        pass
    return -1


def script_scan_batch(offset, carry):
    """Scan one batch of content records for every role name.

    A role is counted only where its name appears in quotes. Unquoted would
    match prose: half these roles are ordinary words (Access, Edit, Delete,
    Support, Finance, Beta) and would hit almost every script.

    CustomAccessRoles.xml is excluded because it lists every role by
    definition, so it would give all 87 a false hit."""
    import json
    offset = int(offset)
    roles = [rn for rn in role_holders().keys() if len(rn) >= 4]
    pats = [(rn, "'" + rn + "'", '"' + rn + '"') for rn in roles]
    found = carry if isinstance(carry, dict) else {}
    n = 0
    try:
        rows = q.QuerySql("""
            SELECT c.Name, c.TypeID, c.Body FROM dbo.Content c WITH (NOLOCK)
            WHERE c.TypeID IN (1,5) AND c.Name NOT IN (%s)
            ORDER BY c.Id OFFSET %d ROWS FETCH NEXT %d ROWS ONLY"""
            % (",".join("'" + x.replace("'", "''") + "'" for x in SCRIPT_SCAN_SKIP),
               offset, SCRIPT_SCAN_BATCH))
        for r in rows:
            n += 1
            body = sval(r, "Body", "") or ""
            if not body:
                continue
            nm = sval(r, "Name", "") or ""
            kind = "python" if sval(r, "TypeID", 0) == 5 else "text"
            for rn, q1, q2 in pats:
                if q1 in body or q2 in body:
                    scan_add(found, rn, nm, kind)
    except Exception as e:
        return None, 0, str(e)[:200]
    return found, n, ""


def script_scan_save(found, scanned, total, secs=0):
    import json
    import datetime as _dt
    payload = {"when": _dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
               "scanned": int(scanned), "total": int(total),
               "secs": int(secs or 0),
               "cap": SCRIPT_SCAN_CAP, "roles": found}
    try:
        model.WriteContentText(SCRIPT_SCAN_CONTENT, safe_json(payload), "")
        return True, ""
    except Exception as e:
        return False, str(e)[:200]


def role_db_usage():
    """What a role gates in THIS church, as opposed to in TouchPoint's code.

    The embedded map comes from TouchPoint's C# and therefore knows nothing
    about your involvements, your reports or your custom roles. FMC and
    PastoralCare do not appear in it at all, yet between them they restrict 76
    involvements. This fills that in, live."""
    out = {}

    def add(role, area, kind, detail, human=""):
        role = (role or "").strip()
        if not role:
            return
        out.setdefault(role, {}).setdefault(area, []).append([kind, detail, human])

    # involvements restricted to a role
    for r in q.QuerySql("""
            SELECT TOP 600 o.LimitToRole, o.OrganizationName, o.OrganizationStatusId
            FROM dbo.Organizations o WITH (NOLOCK)
            WHERE o.LimitToRole > '' ORDER BY o.LimitToRole, o.OrganizationName"""):
        nm = sval(r, "OrganizationName", "") or ""
        if sval(r, "OrganizationStatusId", 30) != 30:
            nm += "  (inactive)"
        add(sval(r, "LimitToRole", ""), "Your involvements", "sees involvement", nm, nm)

    # special content gated to a role
    for r in q.QuerySql("""
            SELECT TOP 400 rl.RoleName, c.Name
            FROM dbo.Content c WITH (NOLOCK)
            JOIN dbo.Roles rl WITH (NOLOCK) ON rl.RoleId = c.RoleID
            WHERE c.RoleID IS NOT NULL ORDER BY rl.RoleName, c.Name"""):
        nm = sval(r, "Name", "") or ""
        add(sval(r, "RoleName", ""), "Your content", "gated content", nm, nm)

    # scripts and reports gated in the CustomReports menu
    body = ""
    try:
        body = model.TextContent("CustomReports") or ""
    except Exception:
        body = ""
    if body:
        import re as _re
        kind_of = {"pyscript": "runs python script", "sqlreport": "runs SQL report",
                   "orgsearchsqlreport": "runs involvement report", "url": "opens page"}
        for tag in _re.findall(r'<Report\b[^>]*?/?>', body):
            nm = _re.search(r'name="([^"]*)"', tag, _re.I)
            ty = _re.search(r'type="([^"]*)"', tag, _re.I)
            ro = _re.search(r'role="([^"]*)"', tag, _re.I)
            if not nm or not ro:
                continue
            k = kind_of.get((ty.group(1) if ty else "").strip().lower(), "runs report")
            for x in ro.group(1).split(","):
                add(x, "Your reports and scripts", k, nm.group(1).strip(), nm.group(1).strip())
    return out


def role_holders():
    out = {}
    for r in q.QuerySql("""
            SELECT r.RoleName, COUNT(ur.UserId) AS N,
                   SUM(CASE WHEN %s IS NULL
                             OR %s < DATEADD(year,-1,GETDATE())
                            THEN 1 ELSE 0 END) AS Dormant
            FROM dbo.Roles r WITH (NOLOCK)
            LEFT JOIN dbo.UserRole ur WITH (NOLOCK) ON ur.RoleId = r.RoleId
            LEFT JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = ur.UserId
            GROUP BY r.RoleName ORDER BY r.RoleName""" % (SEEN, SEEN)):
        out[sval(r, "RoleName", "")] = {"n": sval(r, "N", 0) or 0,
                                        "dormant": sval(r, "Dormant", 0) or 0}
    return out


def role_people(rolename):
    rn = (rolename or "").replace("'", "''")
    seen_log = last_activity_by_user()
    out = []
    for r in q.QuerySql("""
            SELECT TOP 500 u.UserId, u.PeopleId, ISNULL(u.Name2, u.Username) AS Nm,
                   u.Username, %s AS LastSeen,
                   (SELECT COUNT(*) FROM dbo.Users u2 WITH (NOLOCK)
                    WHERE u2.PeopleId = u.PeopleId) AS Logins
            FROM dbo.UserRole ur WITH (NOLOCK)
            JOIN dbo.Roles rr WITH (NOLOCK) ON rr.RoleId = ur.RoleId
            JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = ur.UserId
            WHERE rr.RoleName = '%s'
            ORDER BY LastSeen""" % (SEEN, rn)):
        out.append({"pid": sval(r, "PeopleId", 0), "name": sval(r, "Nm", "") or "",
                    "user": sval(r, "Username", "") or "",
                    "last": datestr(newer(sval(r, "LastSeen"), seen_log.get(sval(r, "UserId")))),
                    "logins": sval(r, "Logins", 1)})
    return out


def token_rows():
    out = []
    for r in q.QuerySql("""
            SELECT t.TokenId, u.UserId, u.PeopleId, ISNULL(u.Name2,u.Username) AS Nm,
                   u.Username, %s AS LastSeen, u.IsLockedOut,
                   t.Expiration, t.CreatedDate
            FROM dbo.UserTokens t WITH (NOLOCK)
            JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = t.UserId
            ORDER BY ISNULL(u.Name2,u.Username), t.CreatedDate""" % SEEN):
        out.append({"id": sval(r, "TokenId", 0), "pid": sval(r, "PeopleId", 0),
                    "who": sval(r, "Nm", "") or "", "user": sval(r, "Username", "") or "",
                    "ownerlast": datestr(sval(r, "LastSeen")),
                    "locked": bool(sval(r, "IsLockedOut", False)),
                    "exp": datestr(sval(r, "Expiration")),
                    "made": datestr(sval(r, "CreatedDate"))})
    return out


def api_use():
    """No token carries a last used date, so the nearest honest signal is
    whether the login has called the API at all. ActivityLog is large, so this
    is bounded to a year."""
    out = {}
    try:
        for r in q.QuerySql("""
                SELECT a.UserId, MAX(a.ActivityDate) AS Last, COUNT(*) AS N
                FROM dbo.ActivityLog a WITH (NOLOCK)
                WHERE a.Activity LIKE 'API:%'
                  AND a.ActivityDate >= DATEADD(day,-365,GETDATE())
                  AND a.UserId IS NOT NULL
                GROUP BY a.UserId"""):
            out[sval(r, "UserId", 0)] = {"last": datestr(sval(r, "Last")),
                                         "n": sval(r, "N", 0)}
    except Exception:
        pass
    return out


# ===========================================================================
# Writes. One person at a time, always confirmed, never silent.
# ===========================================================================

def chosen_ids(pid, raw):
    """Task ids the browser asked for, narrowed to ones that really are open and
    really do belong to this person. The list arrives from the client, so it is
    treated as a request rather than as fact."""
    allowed = set(open_task_ids(pid, "both"))
    want = []
    for x in str(raw or "").replace(" ", "").split(","):
        if x.isdigit() and int(x) in allowed:
            want.append(int(x))
    seen, out = set(), []
    for x in want:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def do_reassign(pid, to_pid, ids):
    """model.TaskNoteMassAssign sets AssigneeId and emails per task through
    EmailTaskNoteInfo, which is why the confirm step states the email count."""
    if not ids:
        return {"ok": False, "n": 0, "message": "nothing selected"}
    model.TaskNoteMassAssign(ids, int(to_pid))
    return {"ok": True, "n": len(ids),
            "message": "reassigned %d task note%s, which sends %d notification%s"
                       % (len(ids), "" if len(ids) == 1 else "s",
                          len(ids), "" if len(ids) == 1 else "s")}


def do_complete(pid, ids):
    """Complete task notes without destroying their notes.

    Every route to complete writes the note field unconditionally:

        TasksNotesModel.CompleteTask   taskNote.Notes = note;   // no null check
        TasksNotesModel.MassComplete   taskNote.Notes = note;   // same

    So passing "" blanks it, and 138,499 open task notes here carry real
    pastoral content in that field. The way round it needs no API and no token:
    read what the note already says and hand the same text straight back, so
    the assignment is a no-op for that column while the status, completed by
    and completed date are all set properly."""
    if not ids:
        return {"ok": False, "n": 0, "message": "nothing selected"}

    existing = {}
    for r in q.QuerySql("""
            SELECT tn.TaskNoteId, ISNULL(tn.Notes,'') AS Notes
            FROM dbo.TaskNote tn WITH (NOLOCK)
            WHERE tn.TaskNoteId IN (%s)""" % ",".join(str(int(i)) for i in ids)):
        existing[sval(r, "TaskNoteId")] = sval(r, "Notes", "")

    done, failed, lastmsg = 0, 0, ""
    for tid in ids:
        try:
            model.TaskNoteComplete(tid, existing.get(tid, ""), None)
            done += 1
        except Exception as e:
            failed += 1
            lastmsg = str(e)[:150]
            if failed > 10 and done == 0:
                return {"ok": False, "n": 0,
                        "message": "stopped after %d failures with none completed. %s"
                                   % (failed, lastmsg)}
    msg = ("completed %d task note%s, each keeping the note it already had"
           % (done, "" if done == 1 else "s"))
    if failed:
        msg += ", %d failed: %s" % (failed, lastmsg)
    return {"ok": done > 0, "n": done, "done_ids": ids, "message": msg}


def auth_token():
    """A token to authenticate the one outbound call this tool has to make.

    Revoking is the only thing here that leaves the building: there is no model
    API for deleting a token, so it has to go through the REST endpoint, and
    model.RestPost carries no ambient authentication. Admin governs who can run
    this page; this is only about how the HTTP call identifies itself.

    Rather than demand a setting, use a token the signed in admin already owns.
    Tokens are stored as plain guids, so this can read one, and it acts as that
    same person, so it grants nothing they did not already have. Which one was
    used is reported back, so it is never a silent borrow.

    An explicit TPxi_AccessAudit_PAT setting still wins if one is configured."""
    try:
        v = model.Setting("TPxi_AccessAudit_PAT", "") or ""
        if v.strip():
            return v.strip(), "the TPxi_AccessAudit_PAT setting"
    except Exception:
        pass
    uid = current_user_id()
    if not uid:
        return "", "no signed in user"
    # The token has to belong to a login that holds Admin, because deleting
    # somebody else's token is gated on CanManageAllUserTokens, which is
    # Roles.Admin. A script only ever learns UserPeopleId, never which login is
    # signed in, so picking any token of theirs is not good enough: this person
    # may well have a login with no Admin on it, and its token would 403.
    for r in q.QuerySql("""
            SELECT TOP 1 t.TokenId, t.Token, u.Username
            FROM dbo.UserTokens t WITH (NOLOCK)
            JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = t.UserId
            WHERE u.PeopleId = %d
              AND (t.Expiration IS NULL OR t.Expiration > GETDATE())
              AND EXISTS (SELECT 1 FROM dbo.UserRole ur WITH (NOLOCK)
                          JOIN dbo.Roles rr WITH (NOLOCK) ON rr.RoleId = ur.RoleId
                          WHERE ur.UserId = u.UserId AND rr.RoleName = 'Admin')
            ORDER BY CASE WHEN t.Expiration IS NULL THEN 1 ELSE 0 END, t.Expiration DESC"""
            % int(uid)):
        return (sval(r, "Token", "") or "",
                "your own token #%s on %s" % (sval(r, "TokenId"), sval(r, "Username", "")))
    # Nothing usable. Separate "you have none" from "you have some but not on a
    # login that can do this", because the fix is different for each.
    try:
        r = q.QuerySql("""
            SELECT COUNT(*) AS N FROM dbo.UserTokens t WITH (NOLOCK)
            JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = t.UserId
            WHERE u.PeopleId = %d
              AND (t.Expiration IS NULL OR t.Expiration > GETDATE())""" % int(uid))
        if r and int(sval(r[0], "N", 0) or 0) > 0:
            return "", "tokens but none on an Admin login"
    except Exception:
        pass
    return "", "no token"


def do_archive(pid, ids):
    """Archive, which is the quiet way to clear a queue.

    model.TaskNoteMassArchive just sets IsArchived = true and submits. No email,
    no note rewritten, and nothing destroyed: the task still exists and can be
    unarchived. Because this report counts only Pending and Accepted rows that
    are not archived, archiving also takes them straight out of the worklist."""
    if not ids:
        return {"ok": False, "n": 0, "message": "nothing selected"}
    try:
        model.TaskNoteMassArchive(ids)
    except Exception as e:
        return {"ok": False, "n": 0, "message": "archive failed: %s" % str(e)[:200]}
    return {"ok": True, "n": len(ids), "done_ids": ids,
            "message": "archived %d task note%s, no emails sent"
                       % (len(ids), "" if len(ids) == 1 else "s")}


def do_delete(pid, ids):
    """Delete, which is permanent.

    model.TaskNoteMassDelete runs real DELETE statements against TaskNote,
    TaskNoteKeyword, TaskNoteExtraValue and TaskNoteExtraValueOption. There is
    no undo and no archive copy. It does write the instruction text to
    ActivityLog on the way out, which is the only trace left.

    It sends no email. It also refuses any task note that is the source of
    another one, and returns how many it rejected, which is passed back here
    rather than swallowed."""
    if not ids:
        return {"ok": False, "n": 0, "message": "nothing selected"}
    try:
        rejected = model.TaskNoteMassDelete(ids)
    except Exception as e:
        return {"ok": False, "n": 0, "message": "delete failed: %s" % str(e)[:200]}
    try:
        rejected = int(rejected or 0)
    except Exception:
        rejected = 0
    n = len(ids) - rejected
    msg = "permanently deleted %d task note%s, no emails sent" % (n, "" if n == 1 else "s")
    if rejected:
        msg += (". %d were refused because another task note is derived from them"
                % rejected)
    return {"ok": n > 0, "n": n, "done_ids": ids, "message": msg}


def revoke_state():
    """Whether revoking can work at all, worked out before anything is drawn.

    Cheaper to tell somebody up front than to let them select fifteen tokens
    and find out on the confirm."""
    pat, how = auth_token()
    if pat:
        return {"ok": True, "how": how}
    return {"ok": False, "how": how}


def do_revoke_tokens(raw_ids):
    """Delete access tokens by TokenId.

    The delete endpoint identifies a token by its VALUE, not its id, and this
    report deliberately never puts a value in the page. That stays true: the
    browser sends ids, the value is read here, used once, and never returned.

    The credential this call makes itself has to belong to a login holding
    Admin. TouchPoint only lets you delete somebody else's token if
    CanManageAllUserTokens passes for the user the TOKEN belongs to, not for
    whoever is signed in to the page, and that check is Roles.Admin.

    Success is confirmed against the database, not against the HTTP call.
    model.RestPost ends with "return response.Content" and never raises, so a
    401, a 403 or a 404 comes back looking exactly like a success. Counting
    those as done would tell an admin a token was revoked while it was still
    live, which is the worst thing this tool could do. So the rows are read
    back from dbo.UserTokens afterwards and only the ones that really went are
    reported gone."""
    ids = []
    for x in str(raw_ids or "").replace(" ", "").split(","):
        if x.isdigit():
            ids.append(int(x))
    if not ids:
        return {"ok": False, "message": "nothing selected"}

    pat, how = auth_token()
    if not pat:
        if how == "tokens but none on an Admin login":
            why = ("You have access tokens, but none of them is on a login that holds "
                   "Admin. Deleting somebody else's token is gated on the role of the "
                   "login the token belongs to, not on who is signed in here. ")
        else:
            why = ("None of your logins has an access token. ")
        return {"ok": False, "message": why +
                "TouchPoint exposes no model API for deleting a token, so this has to call "
                "/v1/Account/DeleteUserAccessToken, and that call needs a credential of its "
                "own whoever is signed in. Create one under Account, Manage Tokens on a "
                "login with Admin, or set TPxi_AccessAudit_PAT in Admin, Settings."}

    host = ""
    try:
        host = model.CmsHost.replace("https://", "").replace("http://", "").strip("/")
    except Exception:
        host = ""
    headers = {"Authorization": "PAT " + pat, "CmsHost": host,
               "Content-Type": "text/plain"}

    idlist = ",".join(str(i) for i in ids)
    vals, before = {}, set()
    for r in q.QuerySql("""
            SELECT t.TokenId, t.Token FROM dbo.UserTokens t WITH (NOLOCK)
            WHERE t.TokenId IN (%s)""" % idlist):
        tid = sval(r, "TokenId")
        vals[tid] = sval(r, "Token", "")
        before.add(tid)

    skipped_self, tried, replies = [], [], []
    for tid in ids:
        v = vals.get(tid)
        if not v:
            continue                      # already gone, handled in the readback
        if v == pat:
            skipped_self.append(tid)      # never cut the branch we are sitting on
            continue
        tried.append(tid)
        try:
            body = model.RestPost(
                "https://api.tpsdb.com/api/v1/Account/DeleteUserAccessToken", headers, v)
            # 204 means an empty body. Anything else is the endpoint talking back.
            if body and str(body).strip():
                replies.append(str(body).strip()[:160])
        except Exception as e:
            replies.append(str(e)[:160])

    # the only answer that counts
    still = set()
    try:
        for r in q.QuerySql("""
                SELECT t.TokenId FROM dbo.UserTokens t WITH (NOLOCK)
                WHERE t.TokenId IN (%s)""" % idlist):
            still.add(sval(r, "TokenId"))
    except Exception:
        return {"ok": False, "message":
                "The calls were made but the check afterwards failed, so this cannot say "
                "which tokens actually went. Reload the page and look at the list."}

    gone = [t for t in ids if t not in still]
    failed = [t for t in tried if t in still]

    bits = []
    if gone:
        bits.append("revoked %d token%s using %s"
                    % (len(gone), "" if len(gone) == 1 else "s", how))
    if failed:
        bits.append("%d did NOT go and is still live" % len(failed)
                    if len(failed) == 1 else
                    "%d did NOT go and are still live" % len(failed))
        if replies:
            bits.append("the server said: " + replies[0])
        else:
            bits.append("the server accepted the call and changed nothing, which usually "
                        "means the token it authenticated with is not allowed to delete "
                        "other people's tokens")
    if skipped_self:
        bits.append("token %s is the one this tool authenticated with, so it was left "
                    "alone. Revoke that one from Manage Tokens"
                    % ", ".join(str(t) for t in skipped_self))
    if not bits:
        bits.append("nothing to do, those tokens were already gone")

    return {"ok": len(gone) > 0, "n": len(gone), "gone": gone,
            "message": ". ".join(bits) + "."}



def login_count(pid):
    """How many logins this person owns. do_remove_role refuses to run unless
    the answer is exactly one, so this has to be right rather than roughly
    right: it is the guard, not a display value."""
    try:
        r = q.QuerySql("""
            SELECT COUNT(*) AS N FROM dbo.Users u WITH (NOLOCK)
            WHERE u.PeopleId = %d""" % int(pid))
        return int(sval(r[0], "N", 0) or 0) if r else 0
    except Exception:
        return 0


def do_remove_role(pid, rolenames):
    """Take one or more roles off a person.

    model.RemoveRole is the only role write exposed to a script, and it does
    this:

        var user = p.Users.FirstOrDefault();

    One login. Whichever one comes back first. On somebody with eight logins it
    strips the role from one and leaves the other seven working, which is
    exactly the failure this report exists to point out. So it is only offered
    where the person has a single login, and refused otherwise rather than
    quietly doing a fraction of the job."""
    pid = int(pid)
    n = login_count(pid)
    if n != 1:
        return {"ok": False, "message":
                "This person has %d logins. model.RemoveRole only ever touches the first "
                "one, so removing from here would leave the others holding the role and "
                "look like it had worked. Use Admin, Manage Users and do each login." % n}
    names = [x.strip() for x in str(rolenames or "").split(",") if x.strip()]
    if not names:
        return {"ok": False, "message": "nothing selected"}
    done = []
    for rn in names:
        try:
            model.RemoveRole("PeopleId=%d" % pid, rn)
            done.append(rn)
        except Exception as e:
            return {"ok": False, "message": "stopped at %s: %s" % (rn, str(e)[:160]),
                    "n": len(done)}
    return {"ok": True, "n": len(done),
            "message": "removed %s" % ", ".join(done)}


# ===========================================================================
# Involvements. What is still active, and what looks finished.
#
# The tool can empty an involvement but it cannot close one. There is no model
# method that touches OrganizationStatusId, /api/v1 has nothing for it, and
# model.ExecuteSql throws outside debug. The only code that sets it lives on
# /APIOrg/UpdateOrganization, which is Basic auth plus the Developer role and
# writes eleven fields when you want one. So the last step is a link to the
# involvement, where a human does it in TouchPoint's own screen.
#
# That split is deliberate rather than a limitation to apologise for: dropping
# members is reversible and closing is the decision, so the decision stays with
# a person.
# ===========================================================================

def involvement_rows(progid=0):
    """Every active involvement with the facts you need to judge it.

    Reads LastMeetingDate and FirstMeetingDate straight off Organizations
    rather than aggregating Meetings, because TouchPoint already maintains
    them. Member counts and last attendance do need a roll up, done set based
    in one pass: a correlated subquery per involvement is what made an earlier
    version of this report time out.

    Program comes through the primary division, which is how TouchPoint files
    an involvement. An involvement can sit in several divisions through
    DivOrgs; only the primary one is shown, and that is the one on the record."""
    out = []
    where = "o.OrganizationStatusId = 30"
    try:
        pid = int(progid or 0)
    except Exception:
        pid = 0
    if pid:
        where += " AND d.ProgId = %d" % pid
    try:
        rows = q.QuerySql("""
            WITH mem AS (
                SELECT m.OrganizationId,
                       COUNT(*) AS Mem,
                       MAX(m.LastAttended) AS LastAtt,
                       SUM(CASE WHEN m.MemberTypeId IN (140,310,320) THEN 1 ELSE 0 END) AS Ldr
                FROM dbo.OrganizationMembers m WITH (NOLOCK)
                GROUP BY m.OrganizationId
            ), vol AS (
                -- Volunteer scheduler signups. A scheduler driven involvement
                -- can be busy with people claiming slots and show nothing on
                -- meetings, attendance or enrolment, so without this it reads
                -- as abandoned.
                --
                -- The scheduler is nine tables. It does not need all nine: both
                -- volunteer tables carry their parent team id on the row, which
                -- skips the sub group tables entirely, and TimeSlots is where
                -- OrganizationId lives. Three joins each.
                --
                -- Two tables because there are two ways to volunteer: on a
                -- standing team, and for one specific meeting.
                SELECT x.OrganizationId,
                       MAX(x.DateVolunteered) AS LastVol,
                       COUNT(*) AS Vols
                FROM (
                    SELECT ts.OrganizationId, v.DateVolunteered
                    FROM dbo.TimeSlotTeamSubGroupVolunteers v WITH (NOLOCK)
                    JOIN dbo.TimeSlotTeams tt WITH (NOLOCK)
                         ON tt.TimeSlotTeamId = v.TimeSlotTeamId
                    JOIN dbo.TimeSlots ts WITH (NOLOCK)
                         ON ts.TimeSlotId = tt.TimeSlotId
                    UNION ALL
                    SELECT ts.OrganizationId, v.DateVolunteered
                    FROM dbo.TimeSlotMeetingTeamSubGroupVolunteers v WITH (NOLOCK)
                    JOIN dbo.TimeSlotMeetingTeams mt WITH (NOLOCK)
                         ON mt.TimeSlotMeetingTeamId = v.TimeSlotMeetingTeamId
                    JOIN dbo.TimeSlotMeetings tm WITH (NOLOCK)
                         ON tm.TimeSlotMeetingId = mt.TimeSlotMeetingId
                    JOIN dbo.TimeSlots ts WITH (NOLOCK)
                         ON ts.TimeSlotId = tm.TimeSlotId
                ) x
                WHERE x.DateVolunteered IS NOT NULL
                GROUP BY x.OrganizationId
            ), churn AS (
                -- People joining, leaving or changing member type. This is the
                -- only signal that works on an involvement which never meets.
                -- Without it a mailing list of 4,156 people looks exactly like
                -- a class that finished in 2024; with it, one shows 1,828
                -- changes and the other shows 7.
                SELECT et.OrganizationId,
                       MAX(et.TransactionDate) AS LastChurn,
                       COUNT(*) AS Churn
                FROM dbo.EnrollmentTransaction et WITH (NOLOCK)
                WHERE et.TransactionDate >= DATEADD(year, -3, GETDATE())
                GROUP BY et.OrganizationId
            )
            SELECT TOP 1200 o.OrganizationId AS Oid, o.OrganizationName AS Nm,
                   ISNULL(p.Name, '') AS Prog, ISNULL(d.Name, '') AS Divi,
                   ISNULL(mem.Mem, 0) AS Mem, ISNULL(mem.Ldr, 0) AS Ldr,
                   mem.LastAtt, o.LastMeetingDate AS LastMtg,
                   churn.LastChurn, ISNULL(churn.Churn, 0) AS Churn,
                   vol.LastVol, ISNULL(vol.Vols, 0) AS Vols,
                   o.CreatedDate AS Made,
                   ISNULL(o.RegistrationTypeId, 0) AS RegT,
                   ISNULL(o.LimitToRole, '') AS Role,
                   ISNULL(ot.Description, '') AS OrgType
            FROM dbo.Organizations o WITH (NOLOCK)
            LEFT JOIN dbo.Division d WITH (NOLOCK) ON d.Id = o.DivisionId
            LEFT JOIN dbo.Program p WITH (NOLOCK) ON p.Id = d.ProgId
            LEFT JOIN lookup.OrganizationType ot WITH (NOLOCK) ON ot.Id = o.OrganizationTypeId
            LEFT JOIN mem ON mem.OrganizationId = o.OrganizationId
            LEFT JOIN churn ON churn.OrganizationId = o.OrganizationId
            LEFT JOIN vol ON vol.OrganizationId = o.OrganizationId
            WHERE %s
            ORDER BY o.OrganizationName""" % where)
        for r in rows:
            la = datestr(sval(r, "LastAtt"))
            lm = datestr(sval(r, "LastMtg"))
            out.append({
                "oid": sval(r, "Oid", 0), "name": sval(r, "Nm", "") or "",
                "prog": sval(r, "Prog", "") or "", "div": sval(r, "Divi", "") or "",
                "mem": sval(r, "Mem", 0) or 0, "ldr": sval(r, "Ldr", 0) or 0,
                "lastatt": la, "lastmtg": lm,
                "attdays": days_since(sval(r, "LastAtt")),
                "mtgdays": days_since(sval(r, "LastMtg")),
                "lastchurn": datestr(sval(r, "LastChurn")),
                "churndays": days_since(sval(r, "LastChurn")),
                "churn": sval(r, "Churn", 0) or 0,
                "lastvol": datestr(sval(r, "LastVol")),
                "voldays": days_since(sval(r, "LastVol")),
                "vols": sval(r, "Vols", 0) or 0,
                "made": datestr(sval(r, "Made")),
                "reg": sval(r, "RegT", 0) or 0,
                "role": sval(r, "Role", "") or "",
                "otype": sval(r, "OrgType", "") or "",
            })
            # how long since ANYTHING happened. None means nothing ever did.
            e = out[-1]
            days = [d for d in (e["attdays"], e["mtgdays"], e["churndays"],
                                e["voldays"]) if d is not None]
            e["quiet"] = min(days) if days else None
    except Exception:
        pass
    return out


def involvement_programs():
    out = []
    try:
        for r in q.QuerySql("""
                SELECT p.Id, p.Name, COUNT(o.OrganizationId) AS N
                FROM dbo.Program p WITH (NOLOCK)
                LEFT JOIN dbo.Division d WITH (NOLOCK) ON d.ProgId = p.Id
                LEFT JOIN dbo.Organizations o WITH (NOLOCK)
                     ON o.DivisionId = d.Id AND o.OrganizationStatusId = 30
                GROUP BY p.Id, p.Name
                HAVING COUNT(o.OrganizationId) > 0
                ORDER BY p.Name"""):
            out.append({"id": sval(r, "Id", 0), "name": sval(r, "Name", "") or "",
                        "n": sval(r, "N", 0) or 0})
    except Exception:
        pass
    return out


def involvement_members(oid):
    """Who is in one involvement, so a drop can be chosen rather than bulk."""
    out = []
    try:
        oid = int(oid)
    except Exception:
        return out
    try:
        for r in q.QuerySql("""
                SELECT TOP 500 m.PeopleId AS Pid, ISNULL(pp.Name2, '') AS Nm,
                       ISNULL(mt.Description, '') AS MT,
                       m.EnrollmentDate AS Enr, m.LastAttended AS LastAtt,
                       ISNULL(pp.IsDeceased, 0) AS Dead,
                       ISNULL(pp.ArchivedFlag, 0) AS Arch
                FROM dbo.OrganizationMembers m WITH (NOLOCK)
                JOIN dbo.People pp WITH (NOLOCK) ON pp.PeopleId = m.PeopleId
                LEFT JOIN lookup.MemberType mt WITH (NOLOCK) ON mt.Id = m.MemberTypeId
                WHERE m.OrganizationId = %d
                ORDER BY pp.Name2""" % oid):
            out.append({"pid": sval(r, "Pid", 0), "name": sval(r, "Nm", "") or "",
                        "mt": sval(r, "MT", "") or "",
                        "enr": datestr(sval(r, "Enr")),
                        "last": datestr(sval(r, "LastAtt")),
                        "dead": bool(sval(r, "Dead", 0)),
                        "arch": bool(sval(r, "Arch", 0))})
    except Exception:
        pass
    return out


def do_drop_members(oid, raw_pids):
    """Drop the selected people from one involvement.

    model.DropOrgMember uses .Single(), so it raises rather than shrugging if
    the person is not a member or somehow has two rows. Each is therefore
    dropped on its own and the failures are counted and reported, instead of
    one bad row killing the rest of the batch.

    The result is checked against the database afterwards for the same reason
    the token revoke is: reporting a drop that did not happen is worse than
    reporting a failure."""
    try:
        oid = int(oid)
    except Exception:
        return {"ok": False, "message": "no involvement"}
    pids = []
    for x in str(raw_pids or "").replace(" ", "").split(","):
        if x.isdigit():
            pids.append(int(x))
    if not oid or not pids:
        return {"ok": False, "message": "nothing selected"}

    lastmsg = ""
    for pid in pids:
        try:
            model.DropOrgMember(pid, oid)
        except Exception as e:
            lastmsg = str(e)[:160]

    still = set()
    try:
        for r in q.QuerySql("""
                SELECT m.PeopleId AS Pid FROM dbo.OrganizationMembers m WITH (NOLOCK)
                WHERE m.OrganizationId = %d AND m.PeopleId IN (%s)"""
                % (oid, ",".join(str(p) for p in pids))):
            still.add(sval(r, "Pid", 0))
    except Exception:
        return {"ok": False, "message":
                "The drops were attempted but the check afterwards failed, so this "
                "cannot say who actually came off. Reopen the involvement to look."}

    gone = [p for p in pids if p not in still]
    msg = "dropped %d of %d" % (len(gone), len(pids))
    if still:
        msg += ", %d still on the roster" % len(still)
        if lastmsg:
            msg += " (" + lastmsg + ")"
    return {"ok": len(gone) > 0, "n": len(gone), "gone": gone, "message": msg}


# ===========================================================================
# Page
# ===========================================================================

def emit(text):
    """The single place this script writes to the response. Isolating the one
    python2 print statement here keeps the rest of the file parseable by a
    python3 syntax checker, which is how it gets tested before upload."""
    print text


def esc(s):
    if s is None:
        s = ""
    return (unicode(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def handle_ajax():
    a = str(getattr(model.Data, "aa_action", "") or "")
    if not a:
        return False
    try:
        if a == "person":
            pid = int(getattr(model.Data, "aa_pid", 0) or 0)
            idx, orgs = build_index()
            row = idx.get(pid, {"pid": pid, "name": "PeopleId %d" % pid})
            po = person_orgs(pid, orgs)
            row = dict(row)
            row["loginlist"] = person_logins(pid)
            row["leadmem"] = person_leadmem(pid)
            row["tasklist"] = open_task_rows(pid)
            row.update(po)
            emit(safe_json({"ok": True, "person": row}))

        elif a == "scriptscan":
            import json as _j
            off = int(getattr(model.Data, "aa_off", 0) or 0)
            carry = {}
            raw = str(getattr(model.Data, "aa_carry", "") or "")
            if raw.startswith("{"):
                try:
                    carry = _j.loads(raw)
                except Exception:
                    carry = {}
            force = str(getattr(model.Data, "aa_force", "")) == "1"
            if off == 0 and not force and not script_scan_needed():
                # somebody else already scanned today
                emit(safe_json({"ok": True, "done": True, "skipped": True,
                                "off": 0, "total": 0, "roles": 0, "carry": {},
                                "message": "already scanned within the last %d hours"
                                           % SCRIPT_SCAN_MAX_AGE_HOURS}))
                return True
            found, n, err = script_scan_batch(off, carry)
            if found is None:
                emit(safe_json({"ok": False, "message": err}))
            else:
                total = script_scan_total()
                # A batch shorter than the page size means the end of the
                # table. That is decided by what came back, not by a count
                # that might have failed.
                done = n < SCRIPT_SCAN_BATCH
                saved, serr = (True, "")
                if done:
                    secs = 0
                    try:
                        secs = int(float(getattr(model.Data, "aa_secs", 0) or 0))
                    except Exception:
                        secs = 0
                    saved, serr = script_scan_save(found, off + n, total, secs)
                emit(safe_json({"ok": True, "done": done, "off": off + n,
                                "total": total, "roles": len(found),
                                "carry": found,
                                "message": serr if not saved else ""}))

        elif a == "invmembers":
            emit(safe_json({"ok": True,
                            "people": involvement_members(getattr(model.Data, "aa_oid", 0))}))

        elif a == "dropmembers":
            emit(safe_json(do_drop_members(getattr(model.Data, "aa_oid", 0),
                                           getattr(model.Data, "aa_ids", ""))))

        elif a == "recent":
            pid = int(getattr(model.Data, "aa_pid", 0) or 0)
            days = int(getattr(model.Data, "aa_days", 90) or 90)
            if days not in (30, 90, 365):
                days = 90
            emit(safe_json({"ok": True, "days": days,
                            "changes": person_changes(pid, days),
                            "activity": person_activity(pid, days)}))

        elif a == "rolepeople":
            emit(safe_json({"ok": True,
                             "people": role_people(str(getattr(model.Data, "aa_role", "") or ""))}))

        elif a == "reassign":
            pid = int(getattr(model.Data, "aa_pid", 0) or 0)
            to = int(getattr(model.Data, "aa_to", 0) or 0)
            ids = chosen_ids(pid, getattr(model.Data, "aa_ids", ""))
            if not pid or not to:
                emit(safe_json({"ok": False, "message": "need both a from and a to person"}))
            elif pid == to:
                emit(safe_json({"ok": False, "message": "those are the same person"}))
            else:
                emit(safe_json(do_reassign(pid, to, ids)))

        elif a == "complete":
            pid = int(getattr(model.Data, "aa_pid", 0) or 0)
            ids = chosen_ids(pid, getattr(model.Data, "aa_ids", ""))
            emit(safe_json(do_complete(pid, ids)))

        elif a == "apply_update":
            emit(safe_json(do_apply_update()))

        elif a in ("archive", "delete"):
            pid = int(getattr(model.Data, "aa_pid", 0) or 0)
            ids = chosen_ids(pid, getattr(model.Data, "aa_ids", ""))
            emit(safe_json(do_archive(pid, ids) if a == "archive" else do_delete(pid, ids)))

        elif a == "revoketokens":
            emit(safe_json(do_revoke_tokens(getattr(model.Data, "aa_ids", ""))))

        elif a == "removerole":
            pid = int(getattr(model.Data, "aa_pid", 0) or 0)
            rn = str(getattr(model.Data, "aa_role", "") or "")
            emit(safe_json(do_remove_role(pid, rn) if (pid and rn)
                            else {"ok": False, "message": "need a person and a role"}))
        else:
            emit(safe_json({"ok": False, "message": "unknown action " + a}))
    except Exception:
        emit(safe_json({"ok": False, "message": traceback.format_exc()[-900:]}))
    return True



# ===========================================================================
# What to do. Ordered by how sure the evidence is, not by size.
# ===========================================================================

def build_actions(idx, orgs, toks):
    people = idx.values()

    def idle(v):
        return v["idledays"] is None or v["idledays"] > 365

    def roles_of(v):
        return v["roles"]

    acts = []

    # --- objective -------------------------------------------------------
    gone = [v for v in people if (v["deceased"] or v["archived"])
            and (v["footprint"] or v["tasks_own"] or v["tasks_asg"] or v["roles"])]
    live_rows, inert_rows = [], []
    for v in sorted(gone, key=lambda x: -x["footprint"]):
        why = "deceased" if v["deceased"] else "archived"
        po = person_orgs(v["pid"], orgs)
        for kind, label in (("notify", "on the notification To line"),
                            ("regfrom", "on the confirmation From line")):
            for o in po[kind]:
                row = {"k": o["name"], "v": "%s, %s" % (v["name"], label),
                       "oid": o["oid"], "pid": v["pid"]}
                if o["state"]:
                    row["v"] = "%s %s, but %s" % (v["name"], label, o["state"])
                    inert_rows.append(row)
                else:
                    live_rows.append(row)
        if v["roles"]:
            live_rows.append({"k": v["name"] + "  (" + why + ")",
                              "v": "%d roles still attached" % v["roles"], "pid": v["pid"]})
    acts.append({"sure": 1, "t": "Clear deceased and archived people out of live settings",
                 "n": len(live_rows), "unit": "changes",
                 "why": "These records are marked deceased or archived in TouchPoint and are "
                        "still attached to roles or to involvement settings that fire.",
                 "where": "Each involvement opens on its Registration tab. Roles are on the "
                          "person record.",
                 "rows": live_rows})

    deadtok = {}
    for t in toks:
        if not t.get("apicalls"):
            deadtok.setdefault(t["user"], []).append(t)
    acts.append({"sure": 1, "t": "Revoke tokens on logins that are not calling the API",
                 "n": sum(len(v) for v in deadtok.values()), "unit": "tokens",
                 "why": "These logins have made no API call in a year, so nothing is using "
                        "their tokens. TouchPoint records no last used date per token, so "
                        "this is judged per login, which is safe here because the whole "
                        "login is idle.",
                 "where": "Account, Manage Tokens, signed in as that user.",
                 "rows": [{"k": deadtok[u][0]["who"] + "  " + u,
                           "v": "%d tokens, none used" % len(deadtok[u]),
                           "pid": deadtok[u][0]["pid"]}
                          for u in sorted(deadtok, key=lambda x: -len(deadtok[x]))]})

    noexp = [t for t in toks if not t["exp"]]
    acts.append({"sure": 1, "t": "Replace the tokens that never expire",
                 "n": len(noexp), "unit": "tokens",
                 "why": "A token with no expiry is a permanent credential carrying that "
                        "login's full role set. If one leaks there is no date on which it "
                        "stops working.",
                 "where": "Revoke and re-issue with an expiry.",
                 "rows": [{"k": t["who"] + "  " + t["user"],
                           "v": "token #%s, created %s" % (t["id"], t["made"] or "unknown"),
                           "pid": t["pid"]} for t in noexp]})

    # --- judgement -------------------------------------------------------
    taskgone = [v for v in people
                if not v["service"] and (v["tasks_own"] + v["tasks_asg"]) > 0
                and (v["deceased"] or v["archived"] or idle(v))]
    taskgone.sort(key=lambda v: -(v["tasks_own"] + v["tasks_asg"]))
    acts.append({"sure": 0, "t": "Reassign or close open task notes held by people who have gone",
                 "n": sum(v["tasks_own"] + v["tasks_asg"] for v in taskgone), "unit": "open tasks",
                 "why": "Open, uncompleted task notes owned by or assigned to someone "
                        "deceased, archived, or not seen in over a year. They do not resolve "
                        "themselves and they are not on anyone else's list. Records with no "
                        "birthdate are excluded, because work parked on Admin or Kiosk is "
                        "deliberate.",
                 "where": "Open the person, then use the buttons on their Open task notes "
                          "section. Reassigning emails once per task.",
                 "rows": [{"k": v["name"] + ("  (deceased)" if v["deceased"]
                                             else ("  (archived)" if v["archived"]
                                                   else "  (last seen %s)" % (v["last"] or "never"))),
                           "v": "%d owned, %d assigned" % (v["tasks_own"], v["tasks_asg"]),
                           "pid": v["pid"]} for v in taskgone]})

    staff = [v for v in people if not v["service"] and not v["deceased"]
             and not v["archived"] and v["roles"] and idle(v)]
    staff.sort(key=lambda v: -v["roles"])
    acts.append({"sure": 0, "t": "Check people carrying roles on a login they no longer use",
                 "n": len(staff), "unit": "people",
                 "why": "These hold roles but have not signed in for over a year. Some will "
                        "have left, some just work elsewhere in the system. Volunteers with "
                        "no roles are excluded on purpose, because never signing in is "
                        "normal for them and is not a finding.",
                 "where": "Admin, Manage Users.",
                 "rows": [{"k": v["name"], "v": "%d roles, last seen %s"
                           % (v["roles"], v["last"] or "never"), "pid": v["pid"]}
                          for v in staff]})

    multi = [v for v in people if v["logins"] > 1 and v["roles"]]
    multi.sort(key=lambda v: -v["logins"])
    acts.append({"sure": 0, "t": "Consolidate people who hold roles on several logins",
                 "n": len(multi), "unit": "people",
                 "why": "Roles and tokens attach to a login, not a person, so disabling one "
                        "account leaves the others working. This is the thing most likely to "
                        "be missed when someone leaves.",
                 "where": "Admin, Manage Users.",
                 "rows": [{"k": v["name"], "v": "%d logins, %d roles" % (v["logins"], v["roles"]),
                           "pid": v["pid"]} for v in multi]})

    # --- deliberately not actions ---------------------------------------
    if inert_rows:
        acts.append({"sure": 2, "t": "Same people, on settings that cannot fire",
                     "n": len(inert_rows), "unit": "settings",
                     "why": "These name the same deceased or archived people, but the "
                            "involvement either has no online registration or has the email "
                            "switched off. Nothing is being sent, and with no online "
                            "registration the Messages section does not appear, so there is "
                            "no screen on which to remove the name. Listed so the number "
                            "above is the count you can actually act on.",
                     "where": "Nothing to do unless that involvement is switched back on.",
                     "rows": inert_rows})

    svc = [v for v in people if v["service"] and (v["footprint"] or v["tasks_own"] or v["tasks_asg"])]
    svc.sort(key=lambda v: -(v["footprint"] + v["tasks_own"] + v["tasks_asg"]))
    acts.append({"sure": 2, "t": "Leave the records with no birthdate alone",
                 "n": len(svc), "unit": "records",
                 "why": "A record with no birthdate is usually a department address or a "
                        "service account rather than a person. They sit on notification "
                        "lists and hold work on purpose. An idle login rule would sweep them "
                        "up, so they are named here to stop that. One real person here also "
                        "has no birthdate, so treat it as a hint.",
                 "where": "No action.",
                 "rows": [{"k": v["name"], "v": "%d involvement settings, %d open tasks"
                           % (v["footprint"], v["tasks_own"] + v["tasks_asg"]), "pid": v["pid"]}
                          for v in svc]})
    return acts


CSS = """
<style>
#aa{font:14px/1.5 -apple-system,Segoe UI,Roboto,sans-serif;color:#222;max-width:1260px}
#aa h2{font-size:20px;margin:0 0 2px}
#aa .sub{color:#777;font-size:12px;margin-bottom:12px}
#aa .tabs{border-bottom:1px solid #ddd;margin-bottom:14px}
#aa .tabs button{border:0;background:none;padding:9px 16px;font:inherit;cursor:pointer;
  color:#777;border-bottom:2px solid transparent;margin-bottom:-1px}
#aa .tabs button.on{color:#222;font-weight:600;border-bottom-color:#0f6d72}
#aa .pane{display:none} #aa .pane.on{display:block}
#aa .row2{display:flex;gap:18px;align-items:flex-start}
#aa .left{width:330px;flex:none;border:1px solid #e2e2e2;border-radius:6px;max-height:70vh;
  overflow:auto;background:#fff}
#aa .right{flex:1;border:1px solid #e2e2e2;border-radius:6px;padding:16px;background:#fff;min-height:220px}
#aa .it{display:flex;gap:8px;align-items:center;padding:7px 11px;border-bottom:1px solid #eee;
  cursor:pointer}
#aa .it:hover{background:#f6f6f5}
#aa .it.on{background:#e3f0ef;box-shadow:inset 3px 0 0 #0f6d72}
#aa .it .n{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
#aa .ct{color:#888;font-size:12px;font-family:ui-monospace,Menlo,monospace}
#aa input[type=search],#aa input[type=text]{padding:7px 10px;border:1px solid #ccc;border-radius:5px;
  width:100%;box-sizing:border-box;font:inherit;margin-bottom:8px}
#aa .sec{border:1px solid #e6e6e6;border-radius:6px;margin-bottom:8px}
#aa .sec>summary{padding:8px 11px;cursor:pointer;background:#fafafa;display:flex;gap:9px;
  align-items:center}
#aa .sec .t{font-weight:600;flex:1}
#aa .pill{background:#eee;color:#666;border-radius:99px;padding:1px 8px;font-size:11px;
  font-family:ui-monospace,monospace}
#aa .pill.hot{background:#fbeedb;color:#a8680f}
#aa .in{padding:9px 12px}
#aa .li{display:flex;gap:9px;padding:4px 0;border-bottom:1px dotted #eee;font-size:13px}
#aa .li:last-child{border-bottom:0}
#aa .li .w{color:#888;font-size:11px;min-width:104px;text-transform:uppercase;letter-spacing:.04em}
#aa .bad{color:#9d2f38;font-weight:600}
#aa .warn{background:#fbeedb;border:1px solid #e8d5b5;border-radius:5px;padding:9px 11px;
  font-size:13px;margin-bottom:10px}
#aa .danger{background:#f8e7e8;border:1px solid #e3bcbf;border-radius:5px;padding:11px 13px;
  font-size:13px;margin:10px 0}
#aa .btn{padding:7px 13px;border:1px solid #ccc;border-radius:5px;background:#f4f4f4;
  font:inherit;font-size:13px;cursor:pointer}
#aa .btn.go{background:#0f6d72;border-color:#0f6d72;color:#fff;font-weight:600}
#aa .btn[disabled]{opacity:.5;cursor:not-allowed}
#aa .muted{color:#999} #aa .mono{font-family:ui-monospace,Menlo,monospace;font-size:12px}
#aa .tag{font-size:10px;text-transform:uppercase;letter-spacing:.05em;padding:1px 6px;
  border-radius:3px;background:#eee;color:#777}
#aa .tag.d{background:#f8e7e8;color:#9d2f38}
#aa table{border-collapse:collapse;width:100%;font-size:13px}
#aa th{text-align:left;font-size:11px;text-transform:uppercase;color:#888;padding:6px 9px;
  border-bottom:1px solid #ddd}
#aa td{padding:5px 9px;border-bottom:1px solid #f0f0f0}
#aa .scroll{overflow-x:auto;max-height:340px;overflow-y:auto;border:1px solid #eee;border-radius:5px}
#aa .tfil{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin:6px 0}
#aa .rchip{display:inline-block;font-size:11px;background:#eee;color:#555;border-radius:3px;padding:1px 6px;margin:0 3px 3px 0;cursor:pointer}
#aa .rchip input{margin:0 3px 0 0;vertical-align:-1px}
#aa a.btn{text-decoration:none;display:inline-block}
#aa a.tag{text-decoration:none;color:#0f6d72;border:1px solid #cfd8d8}
#aa .btn.danger-btn{border-color:#d9b3b6;color:#9d2f38}
#aa .btn.danger-btn:hover{background:#f8e7e8}
#aa .tag.mine{background:#e3f0ef;color:#0f6d72;border:1px solid #bcd7d5}
#aa .btn.rf.on{background:#0f6d72;border-color:#0f6d72;color:#fff}
#aa .tag.g{background:#e3f0ef;border-color:#9fc9c6;color:#0f6d72}
#aa .irow{cursor:pointer}
#aa .irow:hover{background:#f6f6f5}
</style>
"""


def render():
    idx, orgs = build_index()
    # 34 of the ~1000 records here have no birthdate, and all but one are plainly
    # not people: Admin, Kiosk, the Email department addresses, eConnect. They
    # hold enormous task counts and sit on notify lists on purpose. Admin alone
    # owns 259,000 open tasks, so without this it is the entire first screen.
    # A missing birthdate is a hint rather than proof, one real person here has
    # none, so this only sorts them last and labels what was actually observed.
    people = sorted(idx.values(),
                    key=lambda v: (1 if v["service"] else 0,
                                   -(v["footprint"] + v["tasks_own"] + v["tasks_asg"]),
                                   v["name"]))
    holders = role_holders()
    rmap = role_map()
    toks = token_rows()
    use = api_use()

    # attach api activity to each token by its login
    uid_use = {}
    for r in q.QuerySql("SELECT UserId, Username FROM dbo.Users WITH (NOLOCK) WHERE PeopleId IS NOT NULL"):
        uid_use[sval(r, "Username", "")] = use.get(sval(r, "UserId", 0), {"last": None, "n": 0})
    for t in toks:
        u = uid_use.get(t["user"], {"last": None, "n": 0})
        t["apilast"] = u["last"]
        t["apicalls"] = u["n"]

    dbuse = role_db_usage()
    # what the last script scan found, folded in as another church-side area.
    # Items match what role_db_usage.add builds: [kind, detail, human]
    sscan = script_scan_cache()
    if sscan:
        for rn, items in (sscan.get("roles") or {}).items():
            bucket = dbuse.setdefault(rn, {}).setdefault("Your scripts", [])
            for it in items:
                nm = it[0] if it else ""
                kind = it[1] if len(it) > 1 else "script"
                extra = it[2] if len(it) > 2 else 0
                label = nm
                if extra:
                    label += "  (+%d older cop%s)" % (extra, "y" if extra == 1 else "ies")
                bucket.append(["named in " + kind, nm, label])
    sorder, svals = role_screen_settings()
    scat = setting_catalog()
    spos = {}
    for i, rn in enumerate(sorder):
        spos[rn] = i
    roles = []
    for name in sorted(holders.keys()):
        h = holders[name]
        entry = {"n": name, "held": h["n"], "dormant": h["dormant"],
                 "custom": name not in BUILTIN_ROLES}
        # TouchPoint's own screen settings for this role, with its own wording
        vals = svals.get(name)
        if vals is not None:
            sets, nchanged = [], 0
            for c in scat:
                key = c.get("n")
                if key not in vals:
                    continue
                active = vals[key]
                changed = active != bool(c.get("d"))
                if changed:
                    nchanged += 1
                shown = (not active) if c.get("r") else active
                sets.append({"g": c.get("g"), "f": c.get("f"), "t": c.get("t"),
                             "lab": c.get("y") if shown else c.get("no"),
                             "on": bool(shown), "chg": changed, "k": key})
            entry["sets"] = sets
            entry["nchg"] = nchanged
            entry["pos"] = spos.get(name)
        entry["norder"] = len(sorder)
        # Kept apart rather than merged. They answer different questions and are
        # worth different amounts of attention: TouchPoint's half is a fixed
        # list that only changes on a release, this church's half is live and
        # can run to hundreds of rows. Edit gates 213 content records here.
        tp = {}
        if rmap and name in rmap:
            for a, v in rmap[name].items():
                tp[a] = list(v)
        mine = {}
        for a, v in dbuse.get(name, {}).items():
            mine.setdefault(a, [])
            mine[a].extend(v)
        if tp:
            entry["tp"] = tp
        if mine:
            entry["mine"] = mine
        roles.append(entry)

    actions = build_actions(idx, orgs, toks)

    data = {"people": people, "roles": roles, "tokens": toks,
            "actions": actions,
            "hasmap": bool(rmap), "version": APP_VERSION,
            "revoke": revoke_state(),
            "sscan": ({"when": sscan.get("when"), "scanned": sscan.get("scanned"),
                       "total": sscan.get("total"), "secs": sscan.get("secs") or 0,
                       "roles": len(sscan.get("roles") or {}),
                       "due": script_scan_needed()}
                      if sscan else None),
            "sscandue": script_scan_needed(),
            "sscansize": script_scan_total(),
            "invs": involvement_rows(),
            "me": current_user_id() or 0}

    model.Header = "Access Audit"
    model.Form = (CSS + '<div id="aa">'
        + '<div id="appUpdateBanner" style="display:none;background:#e8f0fd;'
        +   'border:1px solid #cfe3ff;border-radius:8px;padding:10px 14px;'
        +   'margin-bottom:12px;align-items:center;gap:10px"></div>'
        + '<h2>Access Audit</h2>'
        + '<div class="sub">Live, from this database. Version ' + APP_VERSION
        + ('' if rmap else
           ' &middot; the role map failed to parse, so By role shows holders only')
        + '</div>'
        + '<div class="tabs">'
        + '<button data-tab="role" class="on">By role</button>'
        + '<button data-tab="person">By person</button>'
        + '<button data-tab="inv">Involvements</button>'
        + '<button data-tab="token">Tokens</button>'
        + '<button data-tab="todo">What to do</button></div>'
        + '<div class="pane" id="pane-person"><div class="row2">'
        +   '<div style="width:330px;flex:none">'
        +     '<input type="search" id="pq" placeholder="Search a person or a login">'
        +     '<div class="left" id="plist"></div></div>'
        +   '<div class="right" id="pdet">Pick a person. Everything they can reach is '
        +     'listed, which is what you work through when someone leaves.</div>'
        + '</div></div>'
        + '<div class="pane on" id="pane-role">'
        +   '<div id="ssbar" class="warn" style="display:none"></div>'
        +   '<div class="row2">'
        +   '<div style="width:330px;flex:none">'
        +     '<input type="search" id="rq" placeholder="Search a role">'
        +     '<div class="tfil" style="margin:0 0 8px">'
        +       '<button class="btn rf on" data-rf="all">All</button> '
        +       '<button class="btn rf" data-rf="builtin">Built in</button> '
        +       '<button class="btn rf" data-rf="custom">Custom</button></div>'
        +     '<div class="left" id="rlist"></div></div>'
        +   '<div class="right" id="rdet">Pick a role.</div>'
        + '</div></div>'
        + '<div class="pane" id="pane-todo">'
        +   '<div class="warn">Nothing here changes anything on its own. The first group '
        +   'is objective, a record says deceased or archived or a login has not called the '
        +   'API in a year. The second needs your judgement. The last group is listed so '
        +   'nobody tidies it away.</div>'
        +   '<div id="alist"></div></div>'
        + '<div class="pane" id="pane-inv">'
        +   '<div class="warn" id="isum"></div>'
        +   '<div class="tfil">'
        +     '<select id="iprog" style="max-width:260px;margin-right:8px"></select>'
        +     '<select id="imon" style="max-width:150px;margin-right:8px">'
        +       '<option value="3">quiet 3+ months</option>'
        +       '<option value="6">quiet 6+ months</option>'
        +       '<option value="12" selected>quiet 12+ months</option>'
        +       '<option value="18">quiet 18+ months</option>'
        +       '<option value="24">quiet 2+ years</option>'
        +       '<option value="36">quiet 3+ years</option></select>'
        +     '<button class="btn ifil on" data-f="quiet">Quiet</button> '
        +     '<button class="btn ifil" data-f="never">Nothing ever</button> '
        +     '<button class="btn ifil" data-f="live">Active</button> '
        +     '<button class="btn ifil" data-f="all">All</button>'
        +     '<input type="search" id="iq" placeholder="Search a name" '
        +       'style="max-width:240px;margin-left:10px">'
        +   '</div>'
        +   '<div class="row2" style="margin-top:10px">'
        +     '<div style="flex:1"><div id="ilist"></div></div>'
        +     '<div class="right" id="idet" style="max-width:520px">'
        +       'Pick an involvement to see who is in it.</div>'
        +   '</div></div>'
        + '<div class="pane" id="pane-token"><div id="tsum" class="warn"></div>'
        +   '<input type="search" id="tq" placeholder="Search an owner or a login">'
        +   '<div id="tlist"></div></div>'
        + '<script>var AA=' + safe_json(data)
        +   ';var APP_VERSION=' + safe_json(APP_VERSION)
        +   ';var DC_SCRIPT_ID=' + safe_json(DC_SCRIPT_ID)
        +   ';var DC_API_BASE=' + safe_json(DC_API_BASE)
        +   ';var SCRIPT_FALLBACK=' + safe_json(get_script_name()) + ';</script>'
        + JS + '</div>')


JS = r"""
<script>
(function(){
var $=function(i){return document.getElementById(i)};
var esc=function(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){
  return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})};
var sel=null, rsel=null, rfilter='all';

function post(params, cb){
  var x=new XMLHttpRequest();
  x.open('POST', window.location.pathname, true);
  x.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
  x.onload=function(){
    if(!x.responseText){ cb({ok:false,message:'empty response from the server'}); return; }
    try{ cb(JSON.parse(x.responseText)); }
    catch(e){ cb({ok:false,message:x.responseText.substring(0,400)}); }
  };
  x.onerror=function(){ cb({ok:false,message:'request failed'}); };
  if(typeof SCRIPT_NAME!=='undefined' && SCRIPT_NAME && !params.script_name){
    params.script_name=SCRIPT_NAME; }
  var b=[]; for(var k in params){ b.push(k+'='+encodeURIComponent(params[k])); }
  x.send(b.join('&'));
}

function idleOf(p){
  if(p.idledays==null) return {t:'no activity recorded', bad:true};
  if(p.idledays>365) return {t:(p.idledays/365).toFixed(1)+' years', bad:true};
  if(p.idledays>90)  return {t:Math.round(p.idledays/30)+' months', bad:false};
  return {t:p.idledays+'d ago', bad:false};
}

/* ---- people ---- */
function drawPeople(){
  var qv=($('pq').value||'').toLowerCase();
  var rows=AA.people.filter(function(p){
    return !qv || p.name.toLowerCase().indexOf(qv)>=0;
  }).slice(0,400);
  $('plist').innerHTML = rows.length ? rows.map(function(p){
    var tags='';
    if(p.deceased) tags+='<span class="tag d">deceased</span>';
    else if(p.archived) tags+='<span class="tag d">archived</span>';
    else if(p.service) tags+='<span class="tag">no birthdate</span>';
    else if(p.logins>1) tags+='<span class="tag">'+p.logins+' logins</span>';
    return '<div class="it'+(sel==p.pid?' on':'')+'" data-pid="'+p.pid+'">'
      +'<span class="n">'+esc(p.name)+'</span>'+tags
      +'<span class="ct">'+(p.footprint+p.tasks_own+p.tasks_asg)+'</span></div>';
  }).join('') : '<div class="in muted">Nobody matches that.</div>';
}

/* Active first, inactive folded away.
   Somebody can be on 140 notification lines and have 128 of them be inactive.
   Mixed together the dozen that still send mail are impossible to pick out,
   and those are the whole point when you are offboarding.

   The inactive ones are kept rather than dropped: an inactive involvement
   sends nothing today, but reactivating it puts that person's name straight
   back on the To line, and by then nobody remembers to check. So they are one
   click away, not gone. */
function orgRows(list,label){
  if(!list||!list.length) return '';
  var row = function(o){
    return '<div class="li"><span class="w">'+label+'</span><span>'+esc(o.name)
      +(o.inactive?' <span class="tag">inactive</span>':'')
      +(o.state?' <span class="tag">'+esc(o.state)+'</span>':'')
      +' <a class="mono" href="/Org/'+o.oid+'#tab-Registrations-tab" target="_blank">open</a>'
      +'</span></div>';
  };
  var live = list.filter(function(o){ return !o.inactive; });
  var dead = list.filter(function(o){ return o.inactive; });
  if(!dead.length) return '<div class="in">'+live.map(row).join('')+'</div>';

  var h = '<div class="in">';
  h += live.length
     ? live.map(row).join('')
     : '<div class="li muted"><span class="w"></span><span>None of these are on an '
       + 'active involvement.</span></div>';
  h += '<details class="sub-inactive" style="margin:6px 0 0">'
    +  '<summary style="cursor:pointer;font-size:12px;color:#777;padding:4px 11px">'
    +  dead.length + ' more on inactive involvements'
    +  '<span class="muted"> &middot; nothing is sent from these unless somebody '
    +  'makes them active again</span></summary>'
    +  dead.map(row).join('')
    +  '</details>';
  return h + '</div>';
}
function sec(title,n,inner,hot,sub){
  return '<details class="sec"'+(n?' open':'')+'><summary><span class="t">'+title+'</span>'
    +(sub?'<span class="muted" style="font-size:12px;margin-left:8px">'+sub+'</span>':'')
    +'<span class="pill'+(n&&hot?' hot':'')+'">'+n+'</span></summary>'
    +(n?inner:'<div class="in muted">Nothing.</div>')+'</details>';
}
// "140" on its own reads as 140 things to deal with. Saying how many are live
// is the difference between a scary number and an accurate one.
function orgSub(list){
  if(!list || !list.length) return '';
  var live = list.filter(function(o){ return !o.inactive; }).length;
  if(live === list.length) return '';
  return live + ' active';
}

function drawPerson(p){
  var id=idleOf(p), tot=(p.tasks_own||0)+(p.tasks_asg||0);
  var h='<h2 style="font-size:18px">'+esc(p.name)+'</h2>'
    +'<div class="sub">PeopleId '+p.pid+' &middot; '+(p.loginlist||[]).length+' logins &middot; '
    +'last active '+(p.last?esc(p.last)+' ('+id.t+')':'no record')+'</div>';
  if(p.deceased||p.archived)
    h+='<div class="danger"><b>This record is marked '+(p.deceased?'deceased':'archived')
      +'.</b> Anything still attached below is live and should come off.</div>';
  else if(p.service)
    h+='<div class="warn"><b>No birthdate on this record.</b> That is usually a '
      +'department address or a service account rather than a person, and work parked '
      +'on those is deliberate. It is a hint, not proof, so check before moving '
      +'anything off it.</div>';
  else if((p.loginlist||[]).length>1)
    h+='<div class="warn"><b>'+p.loginlist.length+' separate logins.</b> Roles and tokens '
      +'attach to a login, not a person, so disabling one leaves the rest working.</div>';

  var one=(p.loginlist||[]).length===1;
  h+=sec('Logins',(p.loginlist||[]).length,'<div class="in">'+(p.loginlist||[]).map(function(l){
      return '<div class="li"><span class="w mono">'+esc(l.user)+'</span><span>'
        +(l.last?esc(l.last):'no activity recorded')
        +(l.locked?' <span class="tag d">locked</span>':'')
        +(l.tokens?' <span class="pill hot">'+l.tokens+' tokens</span>':'')
        +'<br>'+(l.roles.length?l.roles.map(function(r){
            return one
              ? '<label class="rchip"><input type="checkbox" class="rc" value="'+esc(r)
                +'"> '+esc(r)+'</label>'
              : '<span class="tag">'+esc(r)+'</span>'}).join(' ')
            :'<span class="muted">no roles</span>')
        +'</span></div>'}).join('')
      +(one
        ? '<div class="li"><span class="w">roles</span><span>'
          +'<button class="btn" id="rrgo" disabled>Remove selected roles</button> '
          +'<span id="rrmsg" class="muted"></span></span></div>'
        : '<div class="li"><span class="w">roles</span><span>'
          +'<a class="btn" href="/Person2/'+p.pid+'#tab-system" target="_blank">'
          +'Edit roles on the System tab</a> '
          +'<span class="muted">TouchPoint edits roles per login and handles all '
          +p.loginlist.length+' of these properly. A script cannot: the only role write it '
          +'has, <span class="mono">model.RemoveRole</span>, takes just the first login, so '
          +'there is nothing useful to select here.</span></span></div>')
      +'</div>', p.loginlist&&p.loginlist.length>1);

  // hot only when something ACTIVE is attached: an inactive involvement is
  // worth seeing but is not an offboarding blocker
  var liveN = function(l){ return (l||[]).filter(function(o){ return !o.inactive; }).length; };
  h+=sec('Leads these involvements',(p.led||[]).length,orgRows(p.led,'owner'),
         liveN(p.led)>0, orgSub(p.led));
  h+=sec('Leader or assistant in',(p.leadmem||[]).length,
      '<div class="in">'+(p.leadmem||[]).map(function(o){
        return '<div class="li"><span class="w">'+esc(o.mt)+'</span><span>'+esc(o.name)
          +(o.inactive?' <span class="tag">inactive</span>':'')+'</span></div>'}).join('')+'</div>');
  h+=sec('Registration notification, To line',(p.notify||[]).length,
         orgRows(p.notify,'notify'), liveN(p.notify)>5, orgSub(p.notify));
  h+=sec('Registration confirmation, From line',(p.regfrom||[]).length,
         orgRows(p.regfrom,'from'), liveN(p.regfrom)>0, orgSub(p.regfrom));
  h+=sec('Gift notification',(p.giftnotify||[]).length,orgRows(p.giftnotify,'gift'),
         liveN(p.giftnotify)>0, orgSub(p.giftnotify));

  var tl=p.tasklist||[];
  var tasks='<div class="in"><div class="li"><span class="w">owns</span><span>'+(p.tasks_own||0)
    +'</span></div><div class="li"><span class="w">assigned</span><span>'+(p.tasks_asg||0)
    +'</span></div>';
  if(tl.length){
    tasks+='<div class="tfil"><input type="search" id="tkq" placeholder="Filter by person, '
      +'text or keyword" style="max-width:340px;margin:0 8px 0 0">'
      +'<button class="btn" id="tkall">Select all</button> '
      +'<button class="btn" id="tknone">Select none</button>'
      +'<span id="tkn" class="muted" style="margin-left:10px"></span></div>'
      +'<div class="scroll" style="margin:8px 0"><table><thead><tr><th style="width:26px"></th>'
      +'<th>About</th><th>What</th><th>Keyword</th><th>Due</th><th>State</th></tr></thead>'
      +'<tbody id="tkbody">'
      +tl.map(function(t){
        var hay=((t.about||'')+' '+(t.what||'')+' '+(t.kw||'')+' '+(t.org||'')).toLowerCase();
        return '<tr data-hay="'+esc(hay)+'">'
          +'<td><input type="checkbox" class="tkc" value="'+t.id+'"></td>'
          +'<td>'+(t.aboutid?'<a href="/Person2/'+t.aboutid+'" target="_blank">'
                  +esc(t.about||('PeopleId '+t.aboutid))+'</a>':esc(t.about))+'</td>'
          +'<td>'+esc(t.what||'')+(t.org?' <span class="tag">'+esc(t.org)+'</span>':'')+'</td>'
          +'<td class="mono">'+esc(t.kw||'')+'</td>'
          +'<td class="mono">'+esc(t.due||t.made||'')+'</td>'
          +'<td><span class="tag">'+esc(t.status)+'</span> '
          +'<a class="mono" href="/Task/'+t.id+'" target="_blank">#'+t.id+'</a></td></tr>'
      }).join('')+'</tbody></table></div>';
    if(tl.length>=200) tasks+='<div class="muted" style="font-size:12px">Showing the first '
      +'200 of '+tot+'. Only what is listed here can be selected.</div>';
  }
  if(tot){
    tasks+='<div class="muted" style="font-size:12px;padding:2px 0 6px">'
      +'Open means Pending or Accepted and not archived. Completed, archived and '
      +'declined task notes are not counted and cannot be selected.'
      +'</div>'
      +'<div class="danger" id="aa_warn" style="margin:6px 0 4px;display:none"></div>'
      +'<div class="li"><span class="w">move to</span><span>'
      +'<input type="text" id="aa_to" placeholder="PeopleId of the person taking them over" '
      +'style="width:320px;margin:0 6px 0 0">'
      +'<button class="btn go" id="aa_go" disabled>Reassign selected</button> '
      +'<button class="btn" id="aa_done" disabled>Mark selected complete</button> '
      +'<button class="btn" id="aa_arch" disabled>Archive selected</button> '
      +'<button class="btn danger-btn" id="aa_del" disabled>Delete selected</button>'
      +'</span></div>'
      +'<div id="aa_msg" class="muted" style="padding:6px 0"></div>';
  }
  tasks+='</div>';
  h+=sec('Open task notes',tot,tasks,tot>0);

  h+='<details class="sec" id="recsec"><summary><span class="t">What they have been '
    +'doing</span><span class="pill">load</span></summary><div class="in" id="recbody">'
    +'<div class="tfil"><span class="muted" style="margin-right:8px">Last</span>'
    +'<button class="btn rd" data-d="30">30 days</button> '
    +'<button class="btn rd on" data-d="90">90 days</button> '
    +'<button class="btn rd" data-d="365">a year</button></div>'
    +'<div id="recout" class="muted" style="padding:8px 0">Open this to load it. It is '
    +'the one part of this page that is not queried up front.</div></div></details>';
  $('pdet').innerHTML=h;
  wireRecent(p);

  if(tot && tl.length){
    var boxes=function(){ return [].slice.call(document.querySelectorAll('#tkbody .tkc')); };
    var picked=function(){ return boxes().filter(function(c){
        return c.checked && c.closest('tr').style.display!=='none'; }); };
    function refresh(){
      var n=picked().length;
      $('tkn').textContent = n ? n+' selected' : 'nothing selected';
      $('aa_go').disabled = !n; $('aa_done').disabled = !n;
      $('aa_arch').disabled = !n; $('aa_del').disabled = !n;
      $('aa_go').textContent = n ? 'Reassign '+n : 'Reassign selected';
      $('aa_done').textContent = n ? 'Mark '+n+' complete' : 'Mark selected complete';
      $('aa_arch').textContent = n ? 'Archive '+n : 'Archive selected';
      $('aa_del').textContent = n ? 'Delete '+n : 'Delete selected';
      var w=$('aa_warn');
      if(n){ w.style.display='block';
        w.innerHTML='Reassigning and completing each send one email per task note, so either '
          +'would send <b>'+n+' email'+(n===1?'':'s')+'</b>, and there is no way to suppress '
          +'them. <b>Archive</b> sends none and can be undone. <b>Delete</b> sends none and '
          +'cannot.';
      } else { w.style.display='none'; }
    }
    $('tkbody').addEventListener('change',refresh);
    $('tkall').onclick=function(){ boxes().forEach(function(c){
        if(c.closest('tr').style.display!=='none') c.checked=true; }); refresh(); };
    $('tknone').onclick=function(){ boxes().forEach(function(c){ c.checked=false; }); refresh(); };
    $('tkq').oninput=function(){
      var v=this.value.toLowerCase().trim();
      [].slice.call(document.querySelectorAll('#tkbody tr')).forEach(function(tr){
        tr.style.display = (!v || tr.dataset.hay.indexOf(v)>=0) ? '' : 'none';
      });
      refresh();
    };
    function act(kind){
      var ids=picked().map(function(c){ return c.value; });
      if(!ids.length) return;
      var to=($('aa_to').value||'').replace(/\D/g,'');
      var plural=(ids.length===1?'':'s');
      if(kind==='reassign'){
        if(!to){ $('aa_msg').textContent='Enter the PeopleId to move them to.'; return; }
        if(!confirm('Reassign '+ids.length+' task note'+plural+' to PeopleId '+to+'?'
            +'\n\nThis sends '+ids.length+' notification email'+plural+'.')) return;
      } else if(kind==='complete'){
        if(!confirm('Mark '+ids.length+' task note'+plural+' complete?'
            +'\n\nThe existing note on each is left alone, but this sends '+ids.length
            +' email'+plural+'.')) return;
      } else if(kind==='archive'){
        if(!confirm('Archive '+ids.length+' task note'+plural+'?'
            +'\n\nNo emails. They stay in TouchPoint and can be unarchived.')) return;
      } else {
        if(!confirm('PERMANENTLY DELETE '+ids.length+' task note'+plural+'?'
            +'\n\nNo emails, but this cannot be undone. The rows are removed from the '
            +'database, not archived.')) return;
        if(prompt('Type DELETE to confirm.')!=='DELETE'){
          $('aa_msg').textContent='Not confirmed, nothing changed.'; return; }
      }
      ['aa_go','aa_done','aa_arch','aa_del'].forEach(function(b){ $(b).disabled=true; });
      $('aa_msg').textContent='working...';
      post({aa_action:kind, aa_pid:p.pid, aa_to:to, aa_ids:ids.join(',')}, function(r){
        $('aa_msg').textContent=r.message||(r.ok?'done':'failed');
        if(r.ok){ post({aa_action:'person',aa_pid:p.pid}, function(x){
            if(x.ok) drawPerson(x.person); }); }
        else refresh();
      });
    }
    $('aa_go').onclick=function(){ act('reassign'); };
    $('aa_done').onclick=function(){ act('complete'); };
    $('aa_arch').onclick=function(){ act('archive'); };
    $('aa_del').onclick=function(){ act('delete'); };
    refresh();
  }

  if(p.loginlist && p.loginlist.length===1 && $('rrgo')){
    var rboxes=function(){ return [].slice.call(document.querySelectorAll('#pdet .rc')); };
    var rpick=function(){ return rboxes().filter(function(c){ return c.checked; }); };
    function rrefresh(){
      var n=rpick().length;
      $('rrgo').disabled=!n;
      $('rrgo').textContent = n ? 'Remove '+n+' role'+(n===1?'':'s') : 'Remove selected roles';
    }
    $('pdet').addEventListener('change', function(e){
      if(e.target && e.target.classList.contains('rc')) rrefresh(); });
    $('rrgo').onclick=function(){
      var names=rpick().map(function(c){ return c.value; });
      if(!names.length) return;
      if(!confirm('Remove '+names.length+' role'+(names.length===1?'':'s')+' from '
          +p.name+'?\n\n'+names.join(', '))) return;
      $('rrgo').disabled=true; $('rrmsg').textContent='working...';
      post({aa_action:'removerole', aa_pid:p.pid, aa_role:names.join(',')}, function(r){
        $('rrmsg').textContent=r.message||(r.ok?'done':'failed');
        if(r.ok) post({aa_action:'person',aa_pid:p.pid},function(x){
          if(x.ok) drawPerson(x.person); });
      });
    };
    rrefresh();
  }
}

/* ---- roles ---- */
function drawRoles(){
  var qv=($('rq').value||'').toLowerCase();
  var rows=AA.roles.filter(function(r){
    if(rfilter==='custom' && !r.custom) return false;
    if(rfilter==='builtin' && r.custom) return false;
    return !qv||r.n.toLowerCase().indexOf(qv)>=0;});
  $('rlist').innerHTML=rows.map(function(r){
    return '<div class="it'+(rsel==r.n?' on':'')+'" data-role="'+esc(r.n)+'">'
      +'<span class="n">'+esc(r.n)+'</span>'
      +(r.custom?'<span class="tag mine">custom</span>':'')
      +(r.dormant?'<span class="tag">'+r.dormant+' dormant</span>':'')
      +'<span class="ct">'+r.held+'</span></div>'}).join('');
}
/* One area, with a lid on it. Some of these are genuinely long: the Edit role
   gates 213 special content records here, which is a page and a half of
   scrolling past names like "2026 1st RM 211" before you reach anything else.
   So anything over ten rows is collapsed and the rest is behind a button. */
var AREA_CAP = 10;
function areaSec(name, rows){
  var over = rows.length > AREA_CAP;
  var row = function(g){
    return '<div class="li"><span class="w">'+esc(g[0])+'</span><span>'
      +esc(g[2]||g[1])+'</span></div>';
  };
  var inner = '<div class="in">' + rows.slice(0, AREA_CAP).map(row).join('');
  if(over){
    inner += '<div class="amore" style="display:none">'
      + rows.slice(AREA_CAP).map(row).join('') + '</div>'
      + '<div style="padding:6px 0"><button class="btn amt">Show all '
      + rows.length + '</button></div>';
  }
  inner += '</div>';
  // open the short ones, keep the long ones shut so the page stays readable
  return '<details class="sec"'+(rows.length && !over ? ' open' : '')
    +'><summary><span class="t">'+esc(name)+'</span>'
    +'<span class="pill">'+rows.length+'</span></summary>'+inner+'</details>';
}

/* ---- TouchPoint's own screen settings for a role ----
   The wording, the grouping and the on/off labels are all TouchPoint's, read
   out of RoleSettingDefaults.xml. The values are live from CustomAccessRoles.xml.
   Same content as Admin, Roles, Manage, except read only and sitting next to
   everything else the role does.

   Groups stay closed unless something in them was changed, because on a church
   that has never touched these, and that is the normal case, sixty rows of
   "this is the default" is noise in front of the part that matters. */
function screenSettings(r){
  if(!r.sets || !r.sets.length) return '';
  var groups={}, order=[];
  r.sets.forEach(function(x){
    if(!groups[x.g]){ groups[x.g]=[]; order.push(x.g); }
    groups[x.g].push(x);
  });
  var h='<h4 style="margin:16px 0 6px">Screen settings</h4>';
  h+='<div class="'+(r.nchg?'warn':'muted')+'" style="font-size:12px;margin-bottom:8px">'
    +(r.nchg
      ? '<b>'+r.nchg+' of '+r.sets.length+' changed from TouchPoint&rsquo;s default.</b> '
        +'Those are highlighted below.'
      : 'All '+r.sets.length+' are on TouchPoint&rsquo;s defaults. Nobody has customized '
        +'what this role can see.')
    +' These control what appears on screen, tabs, buttons, toolbars, rather than what '
    +'the role unlocks, and they are edited in Admin, Roles.</div>';

  if(r.pos!=null){
    h+='<div class="warn" style="font-size:12px;margin-bottom:8px">'
      +'This role sits at position <b>'+(r.pos+1)+'</b> of '+r.norder+'. For somebody '
      +'holding several roles TouchPoint does not combine these and does not take the most '
      +'permissive: it walks the roles in this order, uses the <b>first</b> one it finds, '
      +'and stops. Every role carries a value for every setting, so whichever of a '
      +'person&rsquo;s roles sits highest decides all '+r.sets.length+' and the rest are '
      +'never read. The order is Roles.Priority with MyData first, and a role with no '
      +'priority set sorts ahead of every role that has one. Adding a role can therefore '
      +'take something away.</div>';
  }

  order.forEach(function(g){
    var rows=groups[g], nch=rows.filter(function(x){return x.chg}).length;
    h+='<details class="sec"'+(nch?' open':'')+'><summary><span class="t">'+esc(g)
      +'</span><span class="pill'+(nch?' hot':'')+'">'+(nch?nch+' changed':rows.length)
      +'</span></summary><div class="in"><div class="scroll"><table><tbody>'
      +rows.map(function(x){
        return '<tr'+(x.chg?' style="background:#fdf6e3"':'')+'>'
          +'<td style="width:110px"><span class="tag'+(x.on?' g':'')+'">'+esc(x.lab)+'</span></td>'
          +'<td style="width:240px">'+esc(x.f)
          +(x.chg?' <span class="tag d">changed</span>':'')+'</td>'
          +'<td class="muted" style="font-size:12px">'+esc(x.t||'')+'</td></tr>';
      }).join('')+'</tbody></table></div></div></details>';
  });
  return h;
}

/* A role's holders, as one bar. Reading "held by 855, 227 not seen in a year"
   tells you less than seeing the shape of it, and the shape is the whole
   question when you are deciding whether a role is over-granted.
   Buckets come from the same last-seen date the table shows, so the bar can
   never disagree with the rows under it. */
function holdersBar(people){
  var now = Date.now(), DAY = 86400000;
  var b = {act:0, mid:0, old:0, never:0};
  people.forEach(function(u){
    if(!u.last){ b.never++; return; }
    var d = (now - new Date(u.last + 'T00:00:00').getTime()) / DAY;
    if(d <= 90)       b.act++;
    else if(d <= 365) b.mid++;
    else              b.old++;
  });
  var n = people.length;
  if(!n) return '';
  var seg = [
    ['act',   b.act,   '#7bc98d', 'active (90d)'],
    ['mid',   b.mid,   '#e0a33e', '3-12 months'],
    ['old',   b.old,   '#e06c78', 'over a year'],
    ['never', b.never, '#adb5b5', 'never signed in']
  ];
  var bar = seg.filter(function(x){ return x[1]; }).map(function(x){
    return '<div title="'+x[1]+' '+esc(x[3])+'" style="width:'
      + (x[1]*100/n) + '%;background:' + x[2] + '"></div>';
  }).join('');
  var key = seg.map(function(x){
    return '<span style="white-space:nowrap;margin-right:14px">'
      + '<span style="display:inline-block;width:9px;height:9px;border-radius:2px;'
      + 'background:'+x[2]+';margin-right:5px"></span>'
      + '<b>'+x[1]+'</b> '+esc(x[3])+'</span>';
  }).join('');
  return '<div style="display:flex;height:9px;border-radius:5px;overflow:hidden;'
    + 'background:#e8eaea;margin:2px 0 7px">' + bar + '</div>'
    + '<div class="muted" style="font-size:12px;margin-bottom:10px">' + key + '</div>';
}

/* ---- scanning your own scripts for role names ----
   TouchPoint's code never checks a custom role, so if one of yours does
   anything at all, a script here is what does it. That is the single most
   useful thing to know about a custom role and the tool could not see it
   until now.

   Runs in batches driven from here rather than in one request, because the
   work scales with how much content a church has, and a big one would time
   out on a single call. */
/* The scan is manual on purpose.
   It was briefly automatic, and that was wrong: 979 records took 4.6 seconds
   on a local copy but 20 seconds against a real server, and that is a small
   church's worth of content. Nobody should wait 20 seconds for a page because
   a background job they did not ask for is walking every script they own.
   So it says what it will cost and waits to be asked. The once-a-day guard
   still lives on the server, so pressing it twice is cheap. */
function ssRender(){
  var b = $('ssbar'), st = AA.sscan;
  if(!b) return;
  var size = (AA.sscansize > 0) ? AA.sscansize : 0;
  var cost = '';
  if(st && st.secs > 0)      cost = 'took ' + st.secs + ' seconds last time';
  else if(size)              cost = 'about ' + size + ' records to read';
  var btn = '<button class="btn" id="ssgo" style="margin-left:6px">'
          + (st && st.when ? 'Scan again' : 'Scan my scripts') + '</button>'
          + (cost ? '<span class="muted" style="margin-left:8px;font-size:12px">'
                    + esc(cost) + '</span>' : '')
          + '<span id="ssmsg" class="muted" style="margin-left:8px"></span>';

  if(st && st.when){
    var scope = (st.total > 0 && st.total >= st.scanned)
      ? (st.scanned + ' of ' + st.total + ' content records')
      : (st.scanned + ' content records');
    b.innerHTML = 'Your scripts were last scanned <b>' + esc(st.when) + '</b>: '
      + scope + ', ' + st.roles + ' roles mentioned.'
      + (st.due ? ' <b>That is over a day old.</b>' : '') + ' ' + btn;
    b.className = st.due ? 'warn' : 'muted';
    b.style.cssText = st.due ? '' : 'font-size:12px;padding:6px 0;margin-bottom:8px';
  } else {
    b.innerHTML = '<b>Your scripts have not been scanned.</b> TouchPoint&rsquo;s own '
      + 'code never checks a custom role, so a script of yours is usually the only '
      + 'thing that does, and this is the only way to find it. It reads your Python '
      + 'and text content, changes nothing, and is worth doing once. ' + btn;
    b.className = 'warn';
    b.style.cssText = '';
  }
  b.style.display = 'block';
  $('ssgo').onclick = function(){ ssRun(true); };
}

function ssRun(forced){
  var btn = $('ssgo'), msg = $('ssmsg'), t0 = Date.now();
  if(btn) btn.disabled = true;
  function say(t){ if(msg) msg.textContent = t; }
  function step(off, carry){
    var pct = (AA.sscansize > 0)
      ? (' (' + Math.min(99, Math.round(off * 100 / AA.sscansize)) + '%)') : '';
    say('scanning ' + off + ' records' + pct + '...');
    post({aa_action:'scriptscan', aa_off:off, aa_force:(forced ? '1' : '0'),
          aa_secs:Math.round((Date.now() - t0) / 1000),
          aa_carry:JSON.stringify(carry||{})},
      function(r){
        if(!r.ok){
          if(msg) msg.innerHTML = '<span class="bad">' + esc(r.message||'failed') + '</span>';
          if(btn) btn.disabled = false;
          return;
        }
        if(r.skipped){ say(r.message || ''); if(btn) btn.disabled = false; return; }
        if(!r.done){ step(r.off, r.carry); return; }
        say('done in ' + Math.round((Date.now() - t0) / 1000) + 's, ' + r.roles
            + ' roles found across ' + r.off + ' records. Reloading.');
        setTimeout(function(){ window.location.reload(); }, 800);
      });
  }
  step(0, {});
}

function drawRole(name){
  var r=null,i;
  for(i=0;i<AA.roles.length;i++){ if(AA.roles[i].n===name){ r=AA.roles[i]; break; } }
  if(!r) return;
  var h='<h2 style="font-size:18px">'+esc(r.n)
    +(r.custom?' <span class="tag mine">custom role</span>'
              :' <span class="tag">built into TouchPoint</span>')+'</h2>'
    +'<div class="sub">held by '+r.held+', '+r.dormant+' of them not seen in a year'
    +(r.custom?'. Created in this database rather than shipped by TouchPoint, by '
               +'whoever set it up. TouchPoint&rsquo;s own code never checks a custom '
               +'role, so anything it unlocks was wired up here.':'')+'</div>';
  if(r.tp||r.mine){
    if(r.tp){
      h+='<h4 style="margin:12px 0 6px">What it unlocks in TouchPoint</h4>'
        +'<div class="muted" style="font-size:12px;margin:-4px 0 8px">Screens, buttons '
        +'and endpoints in TouchPoint&rsquo;s own code. This list only changes when '
        +'TouchPoint ships a release.</div>';
      for(var a in r.tp) h+=areaSec(a, r.tp[a]);
    }
    if(r.mine){
      h+='<h4 style="margin:18px 0 6px">What it gates in your database</h4>'
        +'<div class="muted" style="font-size:12px;margin:-4px 0 8px">Things somebody '
        +'here pointed at this role: involvements, special content, reports and scripts. '
        +'Read live, so it is right as of now.</div>';
      for(var b in r.mine) h+=areaSec(b, r.mine[b]);
    }
  } else if(!AA.hasmap){
    h+='<div class="warn">The built in role map failed to parse, so nothing can be shown '
      +'about what any role unlocks. Everything else on this page is unaffected.</div>';
  } else {
    h+='<div class="warn"><b>Nothing found for this role.</b> Not in TouchPoint&rsquo;s own '
      +'code, not restricting any involvement, not gating any content or report here. '
      +'That does not prove it is unused: it could be checked somewhere this does not '
      +'look, such as inside a python script or a saved query. Treat it as a prompt to '
      +'go and check, never as permission to delete.</div>';
  }
  h+=screenSettings(r);
  h+='<h4 style="margin:14px 0 6px">Who holds it</h4><div id="rp" class="muted">loading...</div>';
  $('rdet').innerHTML=h;
  post({aa_action:'rolepeople',aa_role:name},function(res){
    if(!res.ok){ $('rp').textContent=res.message||'failed'; return; }
    var single=res.people.filter(function(u){ return (u.logins||1)===1; }).length;
    $('rp').innerHTML=
      holdersBar(res.people)
      +'<div class="tfil"><input type="search" id="hq" placeholder="Filter holders" '
      +'style="max-width:280px;margin:0 8px 0 0">'
      +'<button class="btn" id="hall">Select all</button> '
      +'<button class="btn" id="hnone">Select none</button>'
      +'<span id="hn" class="muted" style="margin-left:10px"></span></div>'
      +'<div class="scroll"><table><thead><tr><th style="width:26px"></th><th>Person</th>'
      +'<th>Login</th><th>Last active</th></tr></thead><tbody id="hbody">'
      +res.people.map(function(u){
        var many=(u.logins||1)>1;
        return '<tr data-hay="'+esc((u.name+' '+u.user).toLowerCase())+'">'
          +'<td>'+(many?'<a class="tag" href="/Person2/'+u.pid+'#tab-system" target="_blank" '
                        +'title="'+u.logins+' logins, edit them on the System tab">'+u.logins+'</a>'
                       :'<input type="checkbox" class="hc" value="'+u.pid+'">')+'</td>'
          +'<td><a href="/Person2/'+u.pid+'" target="_blank">'+esc(u.name)+'</a></td>'
          +'<td class="mono">'+esc(u.user)+'</td><td class="mono">'
          +(u.last?esc(u.last):'<span class="bad">never</span>')+'</td></tr>'}).join('')
      +'</tbody></table></div>'
      +'<div class="li" style="margin-top:8px"><span class="w">action</span><span>'
      +'<button class="btn" id="hgo" disabled>Remove this role from selected</button> '
      +'<span id="hmsg" class="muted"></span></span></div>'
      +(single<res.people.length
        ? '<div class="muted" style="font-size:12px;padding:4px 0">'
          +(res.people.length-single)+' of these hold the role on more than one login, shown '
          +'as a numbered link instead of a checkbox. TouchPoint edits roles per login and '
          +'handles that correctly; a script cannot, because '
          +'<span class="mono">model.RemoveRole</span> only touches the first login. The link '
          +'opens their System tab, which is the screen that can.</div>'
        : '');
    var hbox=function(){ return [].slice.call(document.querySelectorAll('#hbody .hc')); };
    var hpick=function(){ return hbox().filter(function(c){
        return c.checked && c.closest('tr').style.display!=='none'; }); };
    function hrefresh(){
      var n=hpick().length;
      $('hn').textContent = n ? n+' selected' : 'nothing selected';
      $('hgo').disabled=!n;
      $('hgo').textContent = n ? 'Remove '+esc(name)+' from '+n : 'Remove this role from selected';
    }
    $('hbody').addEventListener('change', hrefresh);
    $('hall').onclick=function(){ hbox().forEach(function(c){
        if(c.closest('tr').style.display!=='none') c.checked=true; }); hrefresh(); };
    $('hnone').onclick=function(){ hbox().forEach(function(c){ c.checked=false; }); hrefresh(); };
    $('hq').oninput=function(){
      var v=this.value.toLowerCase().trim();
      [].slice.call(document.querySelectorAll('#hbody tr')).forEach(function(tr){
        tr.style.display=(!v||tr.dataset.hay.indexOf(v)>=0)?'':'none'; });
      hrefresh();
    };
    $('hgo').onclick=function(){
      var pids=hpick().map(function(c){ return c.value; });
      if(!pids.length) return;
      if(!confirm('Remove the '+name+' role from '+pids.length+' '
          +(pids.length===1?'person':'people')+'?')) return;
      $('hgo').disabled=true; $('hmsg').textContent='working...';
      var done=0, failed=0, i=0;
      (function step(){
        if(i>=pids.length){
          $('hmsg').textContent='removed from '+done+(failed?(', '+failed+' failed'):'');
          drawRole(name); return;
        }
        post({aa_action:'removerole', aa_pid:pids[i], aa_role:name}, function(r){
          if(r.ok) done++; else failed++;
          i++; step();
        });
      })();
    };
    hrefresh();
  });
}

/* ---- involvements ----
   "Has it met lately" is the obvious test and it is a bad one. Plenty of live
   involvements never meet: Pastor's Daily eConnect has 4,156 people and no
   meeting ever recorded, and by that test it looked as finished as a class
   that ended in 2024.

   What separates them is churn, people joining, leaving or changing member
   type. The mailing list shows 1,828 of those in three years; the finished
   tour shows 7. So "quiet" here is the most recent of three things: the last
   meeting, the last attendance, and the last enrolment change. Nothing in a
   year means nothing at all happened. It is still not a verdict, and the page
   says so. */
var ifil = 'quiet', iprog = 0, isel = null;
var imon = 12;                       // the quiet threshold, in months
var isort = 'quiet', idir = -1;      // column, and 1 asc / -1 desc

function iDays(){ return imon * 30.44; }   // mean month, so 12 lands on a year

function iBucket(r){
  if(r.quiet === null || r.quiet === undefined) return 'never';
  return r.quiet > iDays() ? 'quiet' : 'live';
}

/* Sorting. Nulls always sink to the bottom whichever way the column is
   pointing, because "never attended" is not a small number or a large one and
   letting it sort as 0 would put the emptiest involvements at the top of an
   ascending list and read as though they were the busiest. */
var ICOLS = [
  {k:'name',      t:'Involvement', s:'str'},
  {k:'prog',      t:'Program',     s:'str'},
  {k:'div',       t:'Division',    s:'str'},
  {k:'mem',       t:'People',      s:'num', r:1},
  {k:'quiet',     t:'Quiet for',   s:'num'},
  {k:'lastchurn', t:'Last change', s:'str'},
  {k:'churn',     t:'Churn',       s:'num', r:1}
];
function iCmp(a, b){
  var col = ICOLS.filter(function(c){ return c.k === isort; })[0] || ICOLS[4];
  var x = a[col.k], y = b[col.k];
  var xn = (x === null || x === undefined || x === '');
  var yn = (y === null || y === undefined || y === '');
  if(xn && yn) return 0;
  if(xn) return 1;                       // nulls last, always
  if(yn) return -1;
  if(col.s === 'num') return (x - y) * idir;
  return String(x).toLowerCase() < String(y).toLowerCase() ? -idir : idir;
}
function iRows(){
  var qv = (($('iq')||{}).value || '').toLowerCase();
  return (AA.invs||[]).filter(function(r){
    if(iprog && r.progid !== iprog && r.prog !== iprog) return false;
    if(ifil !== 'all' && iBucket(r) !== ifil) return false;
    if(qv && (r.name + ' ' + r.prog + ' ' + r.div).toLowerCase().indexOf(qv) < 0) return false;
    return true;
  });
}
function drawInvs(){
  var all = AA.invs || [];
  var n = {quiet:0, never:0, live:0};
  all.forEach(function(r){ n[iBucket(r)]++; });
  $('isum').innerHTML =
    '<b>' + all.length + ' active involvements.</b> '
    + n.quiet + ' have had nothing happen in over ' + imon + ' months, ' + n.never
    + ' have nothing recorded at all, ' + n.live + ' saw something this year. '
    + '<span class="muted">Something means a meeting, an attendance, somebody '
    + 'joining, leaving or changing member type, or a volunteer claiming a slot. '
    + 'A list that never meets is still busy if people keep moving through it.</span>'
    + '<br>This tool can empty an involvement but not close one. TouchPoint has no '
    + 'scripting call that sets an involvement inactive, so the last step is the '
    + 'link on each one, which opens it in TouchPoint.';

  var rows = iRows();
  rows.sort(iCmp);
  var shown = rows.slice(0, 400);
  $('ilist').innerHTML =
    '<div class="scroll"><table><thead><tr>'
    + ICOLS.map(function(c){
        var arrow = (isort === c.k) ? (idir === 1 ? ' &uarr;' : ' &darr;') : '';
        return '<th class="isort" data-k="' + c.k + '" style="cursor:pointer'
          + (c.r ? ';text-align:right' : '') + '"'
          + (c.k === 'churn' ? ' title="joins, drops and member type changes in the '
                               + 'last 3 years"' : '')
          + '>' + c.t + arrow + '</th>';
      }).join('')
    + '</tr></thead><tbody>'
    + shown.map(function(r){
        return '<tr class="irow" data-oid="' + r.oid + '"'
          + (isel === r.oid ? ' style="background:#e3f0ef"' : '') + '>'
          + '<td>' + esc(r.name)
          + (r.reg ? ' <span class="tag" title="has a registration">reg</span>' : '')
          + (r.vols ? ' <span class="tag" title="' + r.vols + ' volunteer signups, '
                      + 'last ' + esc(r.lastvol || '?') + '">scheduler</span>' : '')
          + (r.role ? ' <span class="tag d" title="limited to role ' + esc(r.role)
                      + '">' + esc(r.role) + '</span>' : '')
          + '</td>'
          + '<td class="muted">' + esc(r.prog) + '</td>'
          + '<td class="muted">' + esc(r.div) + '</td>'
          + '<td style="text-align:right">' + r.mem
          + (r.ldr ? ' <span class="muted">(' + r.ldr + ' ldr)</span>' : '') + '</td>'
          + '<td class="mono">' + (r.quiet === null || r.quiet === undefined
              ? '<span class="bad">nothing ever</span>'
              : (r.quiet > iDays() ? '<span class="bad">' : '<span>')
                + Math.round(r.quiet / 30) + ' months</span>') + '</td>'
          + '<td class="mono">' + (r.lastchurn ? esc(r.lastchurn)
              : '<span class="muted">none</span>') + '</td>'
          + '<td style="text-align:right" class="mono">' + (r.churn || 0) + '</td>'
          + '</tr>';
      }).join('')
    + '</tbody></table></div>'
    + (rows.length > shown.length
       ? '<div class="muted" style="font-size:12px;padding:6px 0">Showing '
         + shown.length + ' of ' + rows.length + '. Narrow it with the search or the '
         + 'program filter to see the rest.</div>'
       : '');
}
function drawInvDetail(oid){
  isel = oid; drawInvs();
  var r = (AA.invs||[]).filter(function(x){ return x.oid === oid; })[0];
  if(!r) return;
  $('idet').innerHTML = '<h2 style="font-size:16px">' + esc(r.name) + '</h2>'
    + '<div class="sub">' + esc(r.prog) + (r.div ? ' &middot; ' + esc(r.div) : '')
    + (r.otype ? ' &middot; ' + esc(r.otype) : '') + '</div>'
    + '<div class="muted" style="font-size:12px;margin-bottom:8px">'
    + 'created ' + esc(r.made || '?') + ' &middot; last met ' + esc(r.lastmtg || 'never')
    + ' &middot; last attended ' + esc(r.lastatt || 'never')
    + ' &middot; last roster change ' + esc(r.lastchurn || 'never')
    + ' (' + (r.churn || 0) + ' in 3 years)'
    + (r.vols ? ' &middot; last volunteer signup ' + esc(r.lastvol || 'never')
                + ' (' + r.vols + ' total)' : '')
    + '</div>'
    + '<div id="iros" class="muted">loading the roster...</div>';
  post({aa_action:'invmembers', aa_oid:oid}, function(res){
    if(!res.ok){ $('iros').textContent = res.message || 'failed'; return; }
    var p = res.people || [];
    $('iros').innerHTML =
      (p.length
        ? '<div class="tfil"><button class="btn" id="imall">Select all</button> '
          + '<button class="btn" id="imnone">Select none</button>'
          + '<button class="btn danger-btn" id="imgo" disabled>Drop selected</button>'
          + '<span id="immsg" class="muted" style="margin-left:8px"></span></div>'
          + '<div class="scroll" style="margin:8px 0"><table><thead><tr>'
          + '<th style="width:26px"></th><th>Person</th><th>Type</th>'
          + '<th>Enrolled</th><th>Last attended</th></tr></thead><tbody>'
          + p.map(function(u){
              return '<tr><td><input type="checkbox" class="imc" value="' + u.pid + '"></td>'
                + '<td><a href="/Person2/' + u.pid + '" target="_blank">' + esc(u.name) + '</a>'
                + (u.dead ? ' <span class="tag d">deceased</span>' : '')
                + (u.arch ? ' <span class="tag d">archived</span>' : '') + '</td>'
                + '<td class="muted">' + esc(u.mt) + '</td>'
                + '<td class="mono">' + esc(u.enr || '') + '</td>'
                + '<td class="mono">' + esc(u.last || '') + '</td></tr>';
            }).join('') + '</tbody></table></div>'
        : '<div class="muted">Nobody is in this involvement.</div>')
      + '<div class="li" style="margin-top:8px"><span class="w">then</span><span>'
      + '<a class="btn" href="/Org/' + oid + '" target="_blank">Open in TouchPoint</a> '
      + '<span class="muted" style="font-size:12px">Setting an involvement inactive '
      + 'has to be done there. No scripting call exists for it.</span></span></div>';
    if(!p.length) return;
    var boxes = function(){ return [].slice.call(document.querySelectorAll('#iros .imc')); };
    var refresh = function(){
      var n = boxes().filter(function(c){ return c.checked; }).length;
      $('imgo').disabled = !n;
      $('imgo').textContent = n ? ('Drop ' + n) : 'Drop selected';
    };
    $('iros').addEventListener('change', function(e){
      if(e.target && e.target.classList.contains('imc')) refresh(); });
    $('imall').onclick = function(){ boxes().forEach(function(c){ c.checked = true; }); refresh(); };
    $('imnone').onclick = function(){ boxes().forEach(function(c){ c.checked = false; }); refresh(); };
    $('imgo').onclick = function(){
      var ids = boxes().filter(function(c){ return c.checked; }).map(function(c){ return c.value; });
      if(!ids.length) return;
      if(!confirm('Drop ' + ids.length + ' from ' + r.name + '?\n\nThis removes them from '
          + 'the involvement. It does not delete anybody, and they can be added back.')) return;
      $('imgo').disabled = true;
      $('immsg').textContent = 'working...';
      post({aa_action:'dropmembers', aa_oid:oid, aa_ids:ids.join(',')}, function(rr){
        $('immsg').textContent = rr.message || (rr.ok ? 'done' : 'failed');
        if(rr.ok) setTimeout(function(){ drawInvDetail(oid); }, 900);
        else refresh();
      });
    };
  });
}

/* ---- tokens ---- */
function drawTokens(){
  var qv=($('tq').value||'').toLowerCase();
  var noexp=0,far=0,i,t;
  for(i=0;i<AA.tokens.length;i++){ t=AA.tokens[i]; if(!t.exp) noexp++; }
  var rv=AA.revoke||{}, rvnote;
  if(rv.ok){
    rvnote='<br>Revoking is <b>on</b>, using '+esc(rv.how)+'. It takes effect immediately '
      +'and cannot be undone, so anything still using a token stops working the moment it '
      +'goes.';
  } else if(rv.how==='tokens but none on an Admin login'){
    rvnote='<br><b>Revoking is off.</b> You own access tokens, but none of them is on a '
      +'login that holds Admin. TouchPoint decides who may delete somebody else&rsquo;s '
      +'token from the role on the login the token belongs to, not from who is signed in '
      +'here. Create one under Account, Manage Tokens while signed in to an Admin login, '
      +'or set <span class="mono">TPxi_AccessAudit_PAT</span> in Admin, Settings.';
  } else if(rv.how==='no signed in user'){
    rvnote='<br><b>Revoking is off</b> because this page could not tell who is signed in.';
  } else {
    rvnote='<br><b>Revoking is off.</b> It needs an access token of its own, because '
      +'TouchPoint has no scripting call for deleting one. Create one under Account, '
      +'Manage Tokens on a login that holds Admin, or set '
      +'<span class="mono">TPxi_AccessAudit_PAT</span> in Admin, Settings. Everything else '
      +'on this tab works without it.';
  }
  $('tsum').innerHTML='<b>'+AA.tokens.length+' live tokens.</b> '+noexp+' never expire. '
    +'TouchPoint records no last used date on a token, so what is shown per login is '
    +'whether it has called the API in the last year.'+rvnote;
  var by={};
  AA.tokens.forEach(function(t){
    if(qv && (t.who+' '+t.user).toLowerCase().indexOf(qv)<0) return;
    (by[t.user]=by[t.user]||[]).push(t);
  });
  var keys=Object.keys(by).sort(function(a,b){return by[b].length-by[a].length});
  $('tlist').innerHTML=keys.map(function(k){
    var ts=by[k],t0=ts[0],ne=0;
    ts.forEach(function(x){if(!x.exp)ne++});
    return '<details class="sec"><summary><span class="t">'+esc(t0.who)
      +' <span class="mono muted">'+esc(k)+'</span></span>'
      +(t0.apicalls?'<span class="tag">API used '+esc(t0.apilast)+'</span>'
                   :'<span class="tag d">no API calls in a year</span>')
      +'<span class="pill'+(ts.length>20?' hot':'')+'">'+ts.length+'</span></summary>'
      +'<div class="in">'+(ne?'<div class="warn">'+ne+' of these never expire</div>':'')
      +'<div class="tfil"><button class="btn tksel">Select all</button> '
      +'<button class="btn tkclr">Select none</button> '
      +'<button class="btn rvgo" disabled'
      +((AA.revoke&&AA.revoke.ok)?'':' title="Revoking is off, see the note at the top"')
      +'>Revoke selected</button>'
      +'<span class="rvmsg muted" style="margin-left:8px"></span></div>'
      +'<div class="scroll"><table><thead><tr><th style="width:26px"></th><th>Token</th>'
      +'<th>Created</th><th>Expires</th></tr></thead><tbody>'+ts.map(function(x){
        return '<tr><td><input type="checkbox" class="rvc" value="'+x.id+'"></td>'
          +'<td class="mono">#'+x.id+'</td><td class="mono">'+esc(x.made||'')
          +'</td><td class="mono">'+(x.exp?esc(x.exp):'<span class="bad">never</span>')
          +'</td></tr>'}).join('')+'</tbody></table></div></div></details>';
  }).join('');
}

/* ---- wiring ---- */
function showTab(which){
  // every tab button's data-tab has to appear here or the click falls through
  // to the default and lands on By role
  var known={role:1,person:1,inv:1,token:1,todo:1};
  if(!known[which]) which='role';
  document.querySelectorAll('#aa .tabs button').forEach(function(x){
    x.classList.toggle('on', x.dataset.tab===which); });
  document.querySelectorAll('#aa .pane').forEach(function(x){x.classList.remove('on')});
  var el=$('pane-'+which); if(el) el.classList.add('on');
  // keep it in the url so a refresh, or coming back from a link, lands where
  // you were rather than resetting to the first tab
  try{ history.replaceState(null,'','#'+which); }catch(e){ }
}
document.querySelectorAll('#aa .tabs button').forEach(function(b){
  b.onclick=function(){ showTab(b.dataset.tab); };
});
showTab((window.location.hash||'').replace('#','') || 'role');
$('plist').onclick=function(e){
  var d=e.target.closest('.it'); if(!d) return;
  sel=parseInt(d.dataset.pid,10); drawPeople();
  $('pdet').innerHTML='loading...';
  post({aa_action:'person',aa_pid:sel},function(r){
    if(!r.ok){ $('pdet').innerHTML='<div class="danger">'+esc(r.message)+'</div>'; return; }
    drawPerson(r.person);
  });
};
$('rlist').onclick=function(e){
  var d=e.target.closest('.it'); if(!d) return;
  rsel=d.dataset.role; drawRoles(); drawRole(rsel);
};
function drawTodo(){
  var lab={1:['Evidence is objective','d'],0:['Needs your judgement',''],
           2:['Deliberately not an action','']};
  $('alist').innerHTML=(AA.actions||[]).map(function(a){
    var l=lab[a.sure];
    return '<details class="sec"'+((a.sure===1&&a.n)?' open':'')+'>'
      +'<summary><span class="t">'+esc(a.t)+'</span>'
      +'<span class="tag '+l[1]+'">'+l[0]+'</span>'
      +'<span class="pill'+((a.sure===1&&a.n)?' hot':'')+'">'+a.n+' '+a.unit+'</span></summary>'
      +'<div class="in">'
      +'<div class="li"><span class="w">why</span><span>'+a.why+'</span></div>'
      +'<div class="li"><span class="w">where</span><span>'+esc(a.where)+'</span></div>'
      +(a.rows.length?'<div class="scroll" style="margin-top:8px"><table><tbody>'
        +a.rows.map(function(r){
          return '<tr><td>'+(r.pid?'<a href="#" data-goto="'+r.pid+'">'+esc(r.k)+'</a>'
                                  :esc(r.k))+'</td>'
            +'<td class="mono">'+esc(r.v)+'</td>'
            +'<td style="width:56px;text-align:right">'
            +(r.oid?'<a class="mono" target="_blank" href="/Org/'+r.oid
                    +'#tab-Registrations-tab">open</a>':'')+'</td></tr>'}).join('')
        +'</tbody></table></div>':'')
      +'</div></details>'}).join('');
}
$('alist').onclick=function(e){
  var a=e.target.closest('a[data-goto]'); if(!a) return;
  e.preventDefault();
  sel=parseInt(a.dataset.goto,10);
  showTab('person');
  drawPeople(); $('pdet').innerHTML='loading...';
  post({aa_action:'person',aa_pid:sel},function(r){
    if(r.ok) drawPerson(r.person);
    else $('pdet').innerHTML='<div class="danger">'+esc(r.message)+'</div>';
  });
};
function rvBox(d){ return [].slice.call(d.querySelectorAll('.rvc')); }
function rvRefresh(d){
  var n=rvBox(d).filter(function(c){return c.checked}).length;
  var b=d.querySelector('.rvgo');
  // stays dead when there is no credential to revoke with, whatever is ticked
  var can=!!(AA.revoke&&AA.revoke.ok);
  b.disabled = !n || !can;
  b.textContent = (n && can) ? 'Revoke '+n
                : n ? 'Revoking is off' : 'Revoke selected';
}
$('tlist').addEventListener('change',function(e){
  if(e.target && e.target.classList.contains('rvc')) rvRefresh(e.target.closest('details'));
});
$('tlist').addEventListener('click',function(e){
  var t=e.target; if(!t||!t.classList) return;
  var d=t.closest('details'); if(!d) return;
  if(t.classList.contains('tksel')||t.classList.contains('tkclr')){
    var on=t.classList.contains('tksel');
    rvBox(d).forEach(function(c){ c.checked=on; });
    rvRefresh(d);
  } else if(t.classList.contains('rvgo')){
    var ids=rvBox(d).filter(function(c){return c.checked}).map(function(c){return c.value});
    if(!ids.length) return;
    if(!confirm('Revoke '+ids.length+' token'+(ids.length===1?'':'s')+'?\n\n'
        +'Anything still using them stops working immediately. This cannot be undone.')) return;
    t.disabled=true; d.querySelector('.rvmsg').textContent='working...';
    post({aa_action:'revoketokens', aa_ids:ids.join(',')}, function(r){
      var msg=r.message||(r.ok?'done':'failed');
      if(r.gone && r.gone.length){
        // drop them from the data and redraw, rather than reloading the page,
        // which would throw away the tab you are on and the filter you typed
        var dead={}; r.gone.forEach(function(x){ dead[String(x)]=1; });
        AA.tokens=AA.tokens.filter(function(t){ return !dead[String(t.id)]; });
        var keep=($('tq').value||'');
        drawTokens();
        $('tq').value=keep;
        var box=document.querySelector('#tlist .rvmsg');
        if(box) box.textContent=msg;
      } else {
        d.querySelector('.rvmsg').textContent=msg;
        rvRefresh(d);
      }
    });
  }
});
$('rdet').addEventListener('click', function(e){
  var b=e.target;
  if(!b || !b.classList || !b.classList.contains('amt')) return;
  var wrap=b.closest('.in');
  var box=wrap && wrap.querySelector('.amore');
  if(!box) return;
  var shown = box.style.display !== 'none';
  box.style.display = shown ? 'none' : '';
  b.textContent = shown ? ('Show all ' + (box.children.length + AREA_CAP)) : 'Show fewer';
});
document.querySelectorAll('.rf').forEach(function(b){
  b.onclick=function(){
    rfilter=b.dataset.rf;
    document.querySelectorAll('.rf').forEach(function(x){x.classList.toggle('on',x===b)});
    drawRoles();
  };
});

/* ---- what a person has been doing ---- */
function kindTag(k){
  var hot={'role write':1,'delete':1,'merge':1,'giving':1,'download':1};
  return '<span class="tag'+(hot[k]?' d':'')+'">'+esc(k)+'</span>';
}
function drawRecent(r){
  var ch=r.changes||[], ac=r.activity||{}, rows=ac.rows||[], cnt=ac.counts||[];
  var h='';

  h+='<div style="font-weight:600;margin:10px 0 4px">Records they edited</div>';
  if(!ch.length){
    h+='<div class="muted" style="font-size:12px">Nothing in this window that TouchPoint '
      +'could attribute to them. Roughly half of all edits are logged against nobody, '
      +'because a self service edit or a kiosk has no signed in user, so an empty list '
      +'here is not proof they changed nothing.</div>';
  } else {
    var ppl={}; ch.forEach(function(c){ ppl[c.onid]=1; });
    h+='<div class="muted" style="font-size:12px">'+ch.length+' saves across '
      +Object.keys(ppl).length+' people'+(ch.length>=300?', showing the most recent 300':'')
      +'.</div>'
      +'<div class="scroll" style="margin:6px 0"><table><thead><tr><th>When</th>'
      +'<th>Whose record</th><th>Area</th><th>Fields</th></tr></thead><tbody>'
      +ch.map(function(c){
        return '<tr><td class="mono">'+esc(c.when||'')+'</td>'
          +'<td>'+(c.onid?'<a href="/Person2/'+c.onid+'" target="_blank">'+esc(c.on)+'</a>':esc(c.on))
          +(c.self?' <span class="tag">their own</span>':'')+'</td>'
          +'<td>'+esc(c.area||'')+'</td>'
          +'<td class="muted" style="font-size:12px">'+esc(c.flds||'')+'</td></tr>';
      }).join('')+'</tbody></table></div>';
  }

  h+='<div style="font-weight:600;margin:14px 0 4px">What they did in TouchPoint</div>';
  if(!ac.total){
    h+='<div class="muted" style="font-size:12px">No logged activity in this window on any '
      +'of their logins.</div>';
  } else {
    h+='<div class="tfil" style="margin:0 0 6px">'+cnt.map(function(c){
        return kindTag(c.kind)+'<span class="muted" style="margin:0 10px 0 3px">'+c.n+'</span>';
      }).join('')+'</div>';
    if(!rows.length){
      h+='<div class="muted" style="font-size:12px">All '+ac.total+' of those are routine page '
        +'views. Nothing in the kinds worth reading one by one.</div>';
    } else {
      h+='<div class="muted" style="font-size:12px">'+rows.length+' worth reading, out of '
        +ac.total+' logged actions. Routine page views are counted above but not listed.</div>'
        +'<div class="scroll" style="margin:6px 0"><table><thead><tr><th>When</th><th>Kind</th>'
        +'<th>What</th></tr></thead><tbody>'
        +rows.map(function(x){
          return '<tr><td class="mono">'+esc(x.when||'')+'</td><td>'+kindTag(x.kind)+'</td>'
            +'<td>'+esc(x.what||'')
            +(x.onid?' <a class="mono" href="/Person2/'+x.onid+'" target="_blank">#'+x.onid+'</a>':'')
            +'</td></tr>';
        }).join('')+'</tbody></table></div>';
    }
  }
  h+='<div class="muted" style="font-size:12px;margin-top:8px">Read only. Nothing on this '
    +'section changes anything.</div>';
  $('recout').innerHTML=h;
}
function wireRecent(p){
  // named box, not sec: sec is already a global helper in this file
  var box=$('recsec'), loaded={};
  function go(d){
    document.querySelectorAll('#recsec .rd').forEach(function(b){
      b.classList.toggle('on', b.dataset.d===String(d)); });
    if(loaded[d]){ drawRecent(loaded[d]); return; }
    $('recout').innerHTML='<span class="muted">reading the log...</span>';
    post({aa_action:'recent', aa_pid:p.pid, aa_days:d}, function(r){
      if(!r.ok){ $('recout').innerHTML='<span class="bad">'+esc(r.message||'failed')+'</span>';
                 return; }
      loaded[d]=r; drawRecent(r);
    });
  }
  box.addEventListener('toggle', function(){
    if(box.open && !loaded[90]) go(90);
  });
  box.addEventListener('click', function(e){
    if(e.target && e.target.classList && e.target.classList.contains('rd'))
      go(parseInt(e.target.dataset.d,10));
  });
}

/* ---------------- Auto update ---------------- */
var SCRIPT_NAME=(function(){
  try{ var m=(window.location.pathname||'').match(/\/PyScript(?:Form)?\/([^\/?#]+)/);
       if(m&&m[1]) return m[1]; }catch(e){}
  return SCRIPT_FALLBACK;
})();
var APP_LATEST='';
function renderUpdateBanner(){
  var b=$('appUpdateBanner'); if(!b||!APP_LATEST) return;
  b.innerHTML='<div style="font-size:18px">&#128640;</div>'
    +'<div style="flex:1;font-size:12px;color:#0078d4"><strong>Update available</strong>'
    +' &mdash; you have <code>v'+esc(APP_VERSION)+'</code>, latest is <code>v'
    +esc(APP_LATEST)+'</code>. Nothing this tool shows is stored in the script, so '
    +'updating loses nothing.</div>'
    +'<button id="appUpdateBtn" style="white-space:nowrap;padding:6px 14px;'
    +'background:#0078d4;color:#fff;border:0;border-radius:4px;cursor:pointer">'
    +'Update Now</button>';
  b.style.display='flex';
  $('appUpdateBtn').onclick=function(){
    if(!confirm('Update from v'+APP_VERSION+' to v'+APP_LATEST+'?')) return;
    var btn=this; btn.disabled=true; btn.textContent='Updating...';
    post({aa_action:'apply_update', script_name:SCRIPT_NAME}, function(r){
      if(r.ok){ alert(r.message||'Updated'); window.location.reload(true); }
      else { alert('Update failed: '+(r.message||'unknown'));
             btn.disabled=false; btn.textContent='Update Now'; }
    });
  };
}
(function checkForUpdate(){
  try{
    var x=new XMLHttpRequest();
    x.open('GET', DC_API_BASE+'/script-versions', true);
    x.timeout=5000;
    x.onreadystatechange=function(){
      if(x.readyState!==4||x.status!==200) return;
      try{
        var v=JSON.parse(x.responseText)[DC_SCRIPT_ID];
        if(v && v!==APP_VERSION){ APP_LATEST=v; renderUpdateBanner(); }
      }catch(e){}
    };
    x.send();
  }catch(e){}
})();

ssRender();
(function(){
  var seen = {}, opts = ['<option value="0">Every program</option>'];
  (AA.invs||[]).forEach(function(r){ if(r.prog) seen[r.prog] = (seen[r.prog]||0) + 1; });
  Object.keys(seen).sort().forEach(function(k){
    opts.push('<option value="' + esc(k) + '">' + esc(k) + ' (' + seen[k] + ')</option>'); });
  $('iprog').innerHTML = opts.join('');
  $('iprog').onchange = function(){ iprog = this.value === '0' ? 0 : this.value; drawInvs(); };
  $('iq').oninput = drawInvs;
  $('imon').onchange = function(){ imon = parseInt(this.value, 10) || 12; drawInvs(); };
  $('ilist').addEventListener('click', function(e){
    var th = e.target.closest ? e.target.closest('.isort') : null;
    if(!th) return;
    var k = th.dataset.k;
    // same column flips direction; a new column starts on the reading that is
    // useful first, biggest for numbers and A to Z for text
    if(isort === k) idir = -idir;
    else { isort = k; idir = (k === 'name' || k === 'prog' || k === 'div'
                              || k === 'lastchurn') ? 1 : -1; }
    drawInvs();
  });
  document.querySelectorAll('.ifil').forEach(function(b){
    b.onclick = function(){
      ifil = b.dataset.f;
      document.querySelectorAll('.ifil').forEach(function(x){ x.classList.toggle('on', x === b); });
      drawInvs();
    };
  });
  $('ilist').addEventListener('click', function(e){
    var tr = e.target.closest ? e.target.closest('.irow') : null;
    if(tr) drawInvDetail(parseInt(tr.dataset.oid, 10));
  });
  drawInvs();
})();
$('pq').oninput=drawPeople; $('rq').oninput=drawRoles; $('tq').oninput=drawTokens;
drawPeople(); drawRoles(); drawTokens(); drawTodo();
})();
</script>
"""



ROLE_SETTINGS_JSON = r"""[{"n":"CanEditCGInfoEVs","g":"Involvement","f":"Org Extra Values","t":"Allows user to edit existing extra values on an extra value tab","y":"Allow","no":"Disallow","d":false,"r":false},{"n":"DisablePersonLinks","g":"Involvement","f":"Person Links","t":"Enable/Disable hyperlinking to a person's profile (in orgs)","y":"Enable","no":"Disable","d":false,"r":true},{"n":"EditMemberData","g":"Involvement","f":"Edit Member Data","t":"User can edit another user's Org Member Data or not","y":"Yes","no":"No","d":false,"r":false},{"n":"HideExtraValueEdit","g":"Involvement","f":"Org Extra Value Edit Button","t":"Show/Hide the org extra value edit button which allows new extra values to be added","y":"Show","no":"Hide","d":true,"r":false},{"n":"HideGuestsOrgMembers","g":"Involvement","f":"Guest Org Members","t":"Show/Hide Guest Org Member Tab","y":"Show","no":"Hide","d":false,"r":true},{"n":"HideInactiveOrgMembers","g":"Involvement","f":"Inactive Org Members","t":"Show/Hide Inactive Org Member Tab","y":"Show","no":"Hide","d":false,"r":true},{"n":"HidePendingOrgMembers","g":"Involvement","f":"Pending Org Members","t":"Show/Hide Pending Org Member Tab","y":"Show","no":"Hide","d":false,"r":true},{"n":"LeadersCanAlwaysEditOrgContent","g":"Involvement","f":"Edit Org Content","t":"Instead of relying on the EditContent role, Org Leaders can always edit org content","y":"Allow","no":"Disallow","d":false,"r":false},{"n":"LimitedSearchPerson","g":"Involvement","f":"Add New Person Search","t":"Limits search for a name to \"Add Member\" so we don't reveal private info","y":"Limit","no":"Full Access","d":false,"r":false},{"n":"LimitToolbar","g":"Involvement","f":"Blue Toolbar","t":"Custom blue toolbar","y":"Limit","no":"Full Access","d":false,"r":false},{"n":"Organization-CollapseOrgDetails","g":"Involvement","f":"Org Details Box","t":"Expand / Collapse the top Org Details box","y":"Show","no":"Hide","d":false,"r":true},{"n":"Organization-ShowAddress","g":"Involvement","f":"Org Show Address","t":"Show member address","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBirthday","g":"Involvement","f":"Org Show Birthdays","t":"Show member birthday","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbar","g":"Involvement","f":"Show Blue Toolbar","t":"Show blue toolbar","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbarAdminGearMenu","g":"Involvement","f":"Blue Toolbar: Gear Menu","t":"Show blue toolbar: gear menu","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbarCustomReportsMenu","g":"Involvement","f":"Blue Toolbar Menu Option: Custom Reports Menu","t":"Show blue toolbar: custom reports menu","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbarEmailMembers","g":"Involvement","f":"Blue Toolbar Menu Option: Email Members Only","t":"Show blue toolbar: a custom action that emails all the members only","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbarEmailMembersAndProspects","g":"Involvement","f":"Blue Toolbar Menu Option: Email Members and Prospects","t":"Show blue toolbar: a custom action that emails all members AND prospects","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbarEmailProspects","g":"Involvement","f":"Blue Toolbar Menu Option: Email Prospects Only","t":"Show blue toolbar: a custom action that emails prospects only","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbarExportMenu","g":"Involvement","f":"Blue Toolbar: Export Menu","t":"Show blue toolbar: export menu","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbarFullEmailMenu","g":"Involvement","f":"Blue Toolbar: Full Email Menu","t":"Show blue toolbar: full email menu","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbarMembersOnlyPage","g":"Involvement","f":"Blue Toolbar Menu Option: Members Only Page","t":"Add Members only page link to blue toolbar","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowBlueToolbarVolunteerCalendar","g":"Involvement","f":"Blue Toolbar Menu Option: Volunteer Calendar","t":"Add Volunteer calendar link to blue toolbar","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowFiltersBar","g":"Involvement","f":"Filters Bar","t":"Show entire sub-group filters bar","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowOptionsMenu","g":"Involvement","f":"Options Menu","t":"Show Options button and dropdown","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowSettingsTab","g":"Involvement","f":"Settings Tab","t":"Show Settings tab","y":"Show","no":"Hide","d":true,"r":false},{"n":"ShowOrgMembersDropAdd","g":"Involvement","f":"Show Drop/Add Controls","t":"Show/Hide Add Member Dropdown and Drop icon","y":"Show","no":"Hide","d":false,"r":false},{"n":"ShowTagIcon","g":"Involvement","f":"Show/Hide Tag Icon","t":"Show/Hide Tag icon","y":"Show","no":"Hide","d":true,"r":false},{"n":"Meeting-AllowEditDescription","g":"Meetings","f":"Edit Description","t":"Allow editing of meeting 'Description' field","y":"Allow","no":"Disallow","d":true,"r":false},{"n":"Meeting-EnableEditByDefault","g":"Meetings","f":"Enable editing by default","t":"Check the 'Enable - Editing' checkbox by default","y":"On","no":"Off","d":false,"r":false},{"n":"Meeting-HyperlinkNames","g":"Meetings","f":"Hyperlink Names","t":"Enable/Disable hyperlinking to a person's profile (in meetings)","y":"On","no":"Off","d":true,"r":false},{"n":"Meeting-ShowAddGuest","g":"Meetings","f":"Add Guest","t":"Show / Hide 'Add Guest' button","y":"Show","no":"Hide","d":true,"r":false},{"n":"Meeting-ShowAttendType","g":"Meetings","f":"Attend Type","t":"Attendance view: 'Attend Type' column","y":"Show","no":"Hide","d":true,"r":false},{"n":"Meeting-ShowBlueToolbar","g":"Meetings","f":"Blue Toolbar","t":"Show / Hide Meeting blue toolbar","y":"Show","no":"Hide","d":true,"r":false},{"n":"Meeting-ShowBlueToolbarIpadAttendance","g":"Meetings","f":"Blue Toolbar - iPad Attendance","t":"","y":"Show","no":"Hide","d":true,"r":false},{"n":"Meeting-ShowBlueToolbarRollsheet","g":"Meetings","f":"Blue Toolbar - Rollsheet Report","t":"Show / Hide Rollsheet Report option on meetings","y":"Show","no":"Hide","d":true,"r":false},{"n":"Meeting-ShowCurrentMemberType","g":"Meetings","f":"Current Member Type","t":"Attendance view: 'Current Member Type' column","y":"Show","no":"Hide","d":true,"r":false},{"n":"Meeting-ShowEnableBox","g":"Meetings","f":"Enable Box","t":"The 'Enable' set of radio buttons (Editing, Register, Current Members)","y":"Show","no":"Hide","d":true,"r":false},{"n":"Meeting-ShowExtraValuesBox","g":"Meetings","f":"Extra Value Box","t":"Show / Hide Extra Values box","y":"Show","no":"Hide","d":false,"r":false},{"n":"Meeting-ShowOtherAttend","g":"Meetings","f":"Other Attend","t":"Attendance view: 'Other Attend' column","y":"Show","no":"Hide","d":false,"r":false},{"n":"Meeting-ShowShowBox","g":"Meetings","f":"Show Box","t":"The 'Show' set of radio buttons (All, Attends, Absent, Registered)","y":"Show","no":"Hide","d":true,"r":false},{"n":"Meeting-ShowWandTargetBox","g":"Meetings","f":"Wand Target Box","t":"Show/Hide 'Wand Target' box","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowCreateNewMeeting","g":"Meetings","f":"Create New Meeting Button","t":"'Create New Meeting' button","y":"Show","no":"Hide","d":true,"r":false},{"n":"Organization-ShowDeleteMeeting","g":"Meetings","f":"Delete Meeting Button","t":"'Delete Meeting' button","y":"Show","no":"Hide","d":true,"r":false},{"n":"DisableHomePage","g":"General","f":"Home Page","t":"Set home page to user's profile instead of Dashboard","y":"Enable","no":"Disable","d":false,"r":false},{"n":"HideNavTabs","g":"General","f":"Top Navigation Tabs","t":"Show/Hide the main top nav menu bar","y":"Show","no":"Hide","d":false,"r":true},{"n":"OtherGroupsContentOnly","g":"General","f":"Allow Org View to Leaders","t":"Restrict org view to orgs where someone is a leader. For other orgs, he should only see Org Content page","y":"Enable","no":"Disable","d":true,"r":false},{"n":"HideEmailDetails","g":"Person","f":"Email Details","t":"Hides Emails Received, Emails Sent views only show Email Date, From, and Subject. Prevents user from seeing email details which expose search functionality","y":"Limited","no":"Full Access","d":false,"r":false},{"n":"HideMinistryTab","g":"Person","f":"Ministry Tab","t":"Show/Hide Ministry Tab on Person's Profile","y":"Show","no":"Hide","d":false,"r":true},{"n":"HideQueries","g":"Person","f":"Family Member's Link","t":"Disable Family Members link on profile page, because it exposes the search functionality","y":"Enable","no":"Disable","d":false,"r":true},{"n":"Person-ShowBaptism","g":"Person","f":"Show Baptism","t":"Shows or hides the Baptism section on the profile tab","y":"Show","no":"Hide","d":true,"r":false},{"n":"Person-ShowBlueToolbar","g":"Person","f":"Show Blue Toolbar","t":"Show blue toolbar","y":"Show","no":"Hide","d":true,"r":false},{"n":"Person-ShowChurchMembership","g":"Person","f":"Show Church Membership","t":"Shows or hides the Church Membership section on the profile tab","y":"Show","no":"Hide","d":true,"r":false},{"n":"Person-ShowDecision","g":"Person","f":"Show Decision","t":"Shows or hides the Decision section on the profile tab","y":"Show","no":"Hide","d":true,"r":false},{"n":"Person-ShowMemberDocuments","g":"Person","f":"Show Member Documents","t":"Show member documents","y":"Show","no":"Hide","d":true,"r":false},{"n":"Person-ShowDrop","g":"Person","f":"Show Drop","t":"Shows or hides the Drop section on the profile tab","y":"Show","no":"Hide","d":true,"r":false},{"n":"Person-ShowLetterStatus","g":"Person","f":"Show Letter Status","t":"Shows or hides the Letter Status section on the profile tab","y":"Show","no":"Hide","d":true,"r":false},{"n":"Person-ShowMemberProfile","g":"Person","f":"Show Member Profile","t":"Shows or hides the Member Profile label on the profile tab","y":"Show","no":"Hide","d":true,"r":false},{"n":"Person-ShowNewMemberClass","g":"Person","f":"Show New Member Class","t":"Shows or hides the New Member Class section on the profile tab","y":"Show","no":"Hide","d":true,"r":false},{"n":"Person-ShowPledging","g":"Person","f":"Show Pledging Features","t":"Shows or hides the pledge buttons on the giving tab","y":"Show","no":"Hide","d":true,"r":false}]"""

BUILTIN_ROLES = ["Access","AccountManager","Admin","ApiOnly","ApiWrite","AppAdmin","Attendance","BackgroundCheck","CalendarManagement","Checkin","Checkout","ContentEdit","Conversion","Coupon","Coupon2","Delete","Design","Developer","Edit","EditCampus","EmailTemplates","EmailTest","Finance","FinanceAdmin","FinanceDataEntry","FinanceViewOnly","FinanceViewOnlyDetail","FundManager","GivingEmailTemplates","ManageApplication","ManageChat","ManageEmails","ManageEvents","ManageGroups","ManageOrgMembers","ManagePrivacy","ManageProcesses","ManageResources","ManageSMS","ManageTouchpoints","ManageTransactions","Manager","Manager2","McpAccess","McpViewContact","McpViewDemographics","MemberDocs","Membership","MissionGiving","NoRemoteAccess","NotifyLogin","OrgLeadersOnly","OrgTagger","ScheduleEmails","SchedulerTemplates","SendSMS","SpecialContentBasic","SpecialContentFull","StatusFlag","Support","SystemEmailTemplates","TicketScanning","Ticketing","TrainingClasses","UploadPeople","ViewApplication","ViewPrivateTouchpoints","ViewResources","ViewTransactions","VolDocs"]

ROLE_MAP_JSON = """{"Access":{"Core services":[["role check","DirectoryService.cs","Directory"],["role check","PeopleService.cs","People"],["role check","PermissionsService.cs","Permissions"]],"Dialog":[["role check","OrgMemberModel.cs","Org member"],["role check","SearchUsersModel.cs","Search users"]],"ESignature (API)":[["API endpoint","v1/ESignature/ConfigureWebhook","Configure webhook (e signature)"],["API endpoint","v1/ESignature/GetDocument/{requestId:int}","Document (e signature), for a given request"],["API endpoint","v1/ESignature/GetUnfilledSignatureRequests/{orgId:int}","Unfilled signature requests (e signature), for a given org"],["API endpoint","v1/ESignature/SyncSignatureTemplates","Sync signature templates (e signature)"]],"Figures":[["screen element","Index.cshtml","Figures"]],"Giving":[["page or action","GivingManagementController.GetOnlineNotifyPersonList","Online notify person list"],["role check","GivingManagementController.cs","Giving management"]],"Main":[["role check","EmailController.cs","Email"],["screen element","Index.cshtml","Email"],["role check","MassEmailer.cs","Mass emailer"]],"Manage":[["role check","AccountController.cs","Account"]],"Org":[["screen element","Messages.cshtml","Org, Registration, Messages"],["screen element","MessagesEdit.cshtml","Org, Registration, Messages edit"],["role check","VolunteerSchedulerModel.cs","Volunteer scheduler"]],"Other":[["role check","APIFunctions.cs","API functions"],["role check","EngagementScoreHelper.cs","Engagement score helper"],["role check","PeopleSearch.cs","People search"],["role check","Person.cs","Person"],["role check","RealTimeServiceTests.cs","Real time service tests"]],"People":[["screen element","Attendance.cshtml","Person, Enrollment, Attendance"],["screen element","Current.cshtml","Person, Enrollment, Current"],["screen element","Display.cshtml","Person, Profile, Membership, Display"],["screen element","Documents.cshtml","Person, Profile, Membership, Documents"],["screen element","Emails.cshtml","Person, Communications, Emails"],["screen element","Header.cshtml","Person, Personal, Header"],["screen element","Index.cshtml","Person, Touchpoints"],["screen element","Members.cshtml","Person, Family, Members"],["screen element","Pending.cshtml","Person, Enrollment, Pending"],["role check","PersonModel.cs","Person"],["screen element","Previous.cshtml","Person, Enrollment, Previous"],["screen element","Registrations.cshtml","Person, Enrollment, Registrations"],["screen element","Related.cshtml","Person, Family, Related"],["screen element","Tab.cshtml","Person, Profile"],["screen element","TaskNoteIndex.cshtml","Touchpoints, Task note index"],["screen element","Volunteer.cshtml","Person, Enrollment, Volunteer"]],"Permission checks":[["capability","CanAccessTaskNotes","Access task notes"],["capability (also relational)","CanAddPersonTaskNotes","Add person task notes"],["capability (also relational)","CanCheckInFamily","Check in family"],["capability","CanCreateUserToken","Create user token"],["capability","CanManageEmailQueue","Manage email queue"],["capability","CanManageMobileBanners","Manage mobile banners"],["capability (also relational)","CanViewDirectoryPersonProfile","View directory person profile"],["capability (also relational)","CanViewPerson","View person"],["capability (also relational)","CanViewPersonBadges","View person badges"],["capability (also relational)","CanViewPersonChannels","View person channels"],["capability (also relational)","CanViewPersonEmergencyContact","View person emergency contact"],["capability (also relational)","CanViewPersonEngagementScore","View person engagement score"],["capability (also relational)","CanViewPersonInvolvements","View person involvements"],["capability (also relational)","CanViewPersonPrayerRequests","View person prayer requests"],["capability (also relational)","CanViewPersonTaskNotes","View person task notes"]],"Privacy":[["role check","PrivacySettingsModel.cs","Privacy settings"]],"Public":[["page or action","APICheckin2Controller.CheckIn","Check in"],["role check","MobileAPIv2Controller.cs","Mobile AP iv2"]],"Rainforest (API)":[["API endpoint","v1/Rainforest/ConfirmPayin","Confirm payin (rainforest)"],["API endpoint","v1/Rainforest/Payin/{payinId}","Payin (rainforest), for a given payin"],["API endpoint","v1/Rainforest/PaymentMethodConfig","Payment method config (rainforest)"]],"Registrations (API)":[["API endpoint","v1/Involvements/{orgId:int}/Members/{peopleId}/TransactionSummary","Transaction summary (involvements members), for a given org and people"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/Settings","Settings (involvements registrations), for a given org"],["API endpoint","v1/Involvements/{orgId:int}/TransactionSummary","Transaction summary (involvements), for a given org"]],"Setup":[["role check","EmailDelegation.cs","Email delegation"],["role check","LookupController.AccountCodes.cs","Lookup controller account codes"]],"Web":[["screen element","AdminAlerts.cshtml","Shared, Admin alerts"],["screen element","CustomHeader.cshtml","Shared, Custom header"],["screen element","NavBar.cshtml","Shared, Nav bar"],["screen element","People.cshtml","Shared, Menu, People"],["screen element","SearchBox.cshtml","Shared, Search box"]]},"AccountManager":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Org":[["screen element","AccountingCode.cshtml","Org, Editor templates, Accounting code"]],"Permission checks":[["capability","CanManageOrViewLimitedAccountCodes","Manage or view limited account codes"]]},"Admin":{"CheckIn":[["page or action","CheckInDashboardController.CheckInDashboard","Check in dashboard"],["page or action","CheckinLabelsController","Checkin labels controller"],["page or action","ClassroomDashboardController.Index","Index classroom dashboard"],["page or action","ClassroomDashboardController.Person","Person classroom dashboard"],["page or action","FieldRequirementController","Field requirement controller"],["page or action","LabelEditorController","Label editor controller"],["page or action","WebCheckinLabelsController","Web checkin labels controller"]],"ContentSafety (API)":[["API endpoint","v1/ContentSafety/CheckImageContent","Check image content (content safety)"],["API endpoint","v1/ContentSafety/CheckTextContent","Check text content (content safety)"]],"Core services":[["role check","ApplicationSettingsService.cs","Application settings"],["role check","PermissionsService.cs","Permissions"]],"Dialog":[["screen element","Display.cshtml","Org member dialog, Display"],["screen element","Edit.cshtml","Org member dialog, Edit"],["screen element","Entry.cshtml","Org member dialog, Tabs, Extra value, Entry"],["screen element","Groups.cshtml","Org member dialog, Tabs, Groups"],["screen element","History.cshtml","Transaction history, History"],["screen element","Index.cshtml","Org members update"],["screen element","MemberData.cshtml","Org member dialog, Tabs, Member data"],["page or action","OrgMemberDialogController.SmallGroupChecked","Small group checked"],["screen element","Questions.cshtml","Org member dialog, Tabs, Questions"]],"ESignature (API)":[["API endpoint","v1/ESignature/ConfigureWebhook","Configure webhook (e signature)"],["API endpoint","v1/ESignature/GetDocument/{requestId:int}","Document (e signature), for a given request"],["API endpoint","v1/ESignature/GetUnfilledSignatureRequests/{orgId:int}","Unfilled signature requests (e signature), for a given org"],["API endpoint","v1/ESignature/SyncSignatureTemplates","Sync signature templates (e signature)"]],"Finance":[["role check","BundleModel.cs","Bundle"],["role check","FinanceReportsController.cs","Finance reports"],["page or action","FundController","Fund controller"],["page or action","FundController.Index","Index fund"],["page or action","FundSetsController","Fund sets controller"],["screen element","ManagedGiving.cshtml","Finance reports, Managed giving"],["screen element","ManagedPledges.cshtml","Finance reports, Managed pledges"]],"Giving":[["page or action","GivingManagementController.CheckUrlAvailability","Check URL availability"],["page or action","GivingManagementController.Create","Create giving management"],["page or action","GivingManagementController.Delete","Delete giving management"],["page or action","GivingManagementController.GetAvailableFunds","Available funds"],["page or action","GivingManagementController.GetCampusList","Campus list"],["page or action","GivingManagementController.GetConfirmationEmailList","Confirmation email list"],["page or action","GivingManagementController.GetEntryPoints","Entry points"],["page or action","GivingManagementController.GetOnlineNotifyPersonList","Online notify person list"],["page or action","GivingManagementController.GetShellList","Shell list"],["page or action","GivingManagementController.Index","Index giving management"],["page or action","GivingManagementController.List","List giving management"],["page or action","GivingManagementController.Manage","Manage giving management"],["page or action","GivingManagementController.New","New giving management"],["page or action","GivingManagementController.SaveGivingPageEnabled","Save giving page enabled"],["page or action","GivingManagementController.SetGivingDefaultPage","Set giving default page"],["page or action","GivingManagementController.Update","Update giving management"]],"Involvements (API)":[["API endpoint","v1/Involvements","Involvements"],["API endpoint","v1/Involvements/Divisions/Search","Search (involvements divisions)"]],"Main":[["role check","AdminController.cs","Admin"],["screen element","Announcement.cshtml","Announcement, Announcement"],["role check","EmailTemplateModel.cs","Email template"],["screen element","List.cshtml","Coupon, List"],["role check","MassEmailer.cs","Mass emailer"],["page or action","TagsController.ConvertTagToExtraValue","Convert tag to extra value"]],"Manage":[["page or action","BatchMoveAndDeleteController.MoveAndDelete","Move and delete"],["screen element","Details.cshtml","Emails, Details"],["role check","DisplayController.cs","Display"],["page or action","DuplicatesController","Duplicates controller"],["screen element","EditPythonScript.cshtml","Display, Edit python script"],["screen element","EditSqlScript.cshtml","Display, Edit SQL script"],["role check","EmailModel.cs","Email"],["page or action","EmailTemplatesController","Email templates controller"],["role check","EmailsController.cs","Emails"],["role check","EmailsModel.cs","Emails"],["page or action","ExtraValuesController.DeleteAll","Delete all"],["screen element","Index.cshtml","Merge"],["page or action","InvolvementController","Involvement controller"],["screen element","List.cshtml","Org members, List"],["page or action","McpController","Mcp controller"],["page or action","MediaController.CreateResource","Create resource"],["page or action","MediaController.CreateResourceCategory","Create resource category"],["page or action","MediaController.CreateResourceType","Create resource type"],["page or action","MediaController.DeleteAttachment","Delete attachment"],["page or action","MediaController.DeleteResource","Delete resource"],["page or action","MediaController.EditResource","Edit resource"],["page or action","MediaController.MoveResource","Move resource"],["page or action","MediaController.MoveResourceCategoryList","Move resource category list"],["page or action","MediaController.SaveAttachmentListOrder","Save attachment list order"],["page or action","MediaController.SaveResourceListOrder","Save resource list order"],["page or action","MediaController.UpdateAttachment","Update attachment"],["page or action","MediaController.UploadAttachment","Upload attachment"],["page or action","MergeController","Merge controller"],["screen element","OtherDeveloperActions.cshtml","Batch, Other developer actions"],["page or action","ResourceController","Resource controller"],["role check","TransactionsController.cs","Transactions"],["role check","TransactionsModel.cs","Transactions"],["role check","UpdateFieldsModel.cs","Update fields"],["page or action","UsersController","Users controller"]],"MeetingCategory":[["page or action","MeetingCategoryController.Edit","Edit meeting category"]],"Meetings (API)":[["API endpoint","v1/Involvements/{orgId:int}/Meetings","Meetings (involvements), for a given org"]],"OnlineReg":[["screen element","RegPeople.cshtml","Online reg, Other, Reg people"]],"Org":[["screen element","Directory.cshtml","Org, Settings, Directory"],["screen element","Fees.cshtml","Org, Registration, Fees"],["screen element","FeesEdit.cshtml","Org, Registration, Fees edit"],["screen element","Gear.cshtml","Org, Toolbar, Gear"],["role check","General.cs","General"],["screen element","General.cshtml","Org, Settings, General"],["screen element","GeneralEdit.cshtml","Org, Settings, General edit"],["role check","MeetingController.cs","Meeting"],["screen element","MeetingHeader.cshtml","Meeting, Meeting header"],["role check","MemberDirectoryController.cs","Member directory"],["screen element","MessagesEdit.cshtml","Org, Registration, Messages edit"],["role check","OrganizationModel.cs","Organization"],["screen element","Registration.cshtml","Org, Registration, Registration"],["screen element","RegistrationEdit.cshtml","Org, Registration, Registration edit"],["screen element","RegistrationFormRegistration.cshtml","Org, Registration form registration"],["screen element","Results.cshtml","Org search, Results"],["screen element","Settings.cshtml","Org, Settings"],["role check","SettingsAttendanceModel.cs","Settings attendance"],["role check","VolunteerSchedulerModel.cs","Volunteer scheduler"]],"Other":[["role check","Condition.Miscellaneous.cs","Condition miscellaneous"],["role check","DbUtil.Context.cs","Db util context"],["role check","DbUtil.Db.cs","Db util db"],["role check","DbUtil.Session2.cs","Db util session2"],["role check","EngagementScoreHelper.cs","Engagement score helper"],["role check","NavBarPermissions.cs","Nav bar permissions"],["role check","OrganizationMember.cs","Organization member"],["role check","Person.cs","Person"],["role check","PythonScript.cs","Python script"],["role check","ReplacementCodeService_Email.cs","Replacement code service email"],["role check","StandardSettings.cs","Standard settings"],["role check","TasksNotesModel.cs","Tasks notes"],["role check","ToolHelpers.cs","Tool helpers"],["role check","TouchPointAuthenticationMiddleware.cs","Touch point authentication middleware"]],"People":[["role check","AttendHistoryModel.cs","Attend history"],["screen element","Attendance.cshtml","Person, Enrollment, Attendance"],["role check","CommunicationsController.cs","Communications"],["role check","ContactModel.cs","Contact"],["screen element","Display.cshtml","Person, Personal, Display"],["screen element","Edit.cshtml","Person, Personal, Edit"],["role check","EmailModel.cs","Email"],["role check","FailedMailModel.cs","Failed mail"],["screen element","Gear.cshtml","Person, Toolbar, Gear"],["screen element","Index.cshtml","Person"],["role check","LookupsController.cs","Lookups"],["screen element","MessagesNotificationsLog.cshtml","Person, Communications, Messages notifications log"],["role check","PersonController.cs","Person"],["role check","PersonModel.cs","Person"],["screen element","PictureDialog.cshtml","Person, Personal, Picture dialog"],["role check","ProfileController.cs","Profile"],["screen element","Registrations.cshtml","Person, Enrollment, Registrations"],["screen element","UserEdit.cshtml","Person, System, User edit"],["screen element","Users.cshtml","Person, System, Users"]],"People (API)":[["API endpoint","v1/People/Search/AutocompleteEx/Registration","Registration (people search autocomplete ex)"]],"Permission checks":[["capability","CanApproveMeetings","Approve meetings"],["capability","CanApproveReservable","Approve reservable"],["capability","CanChangeTranslations","Change translations"],["capability (also relational)","CanCheckInFamily","Check in family"],["capability","CanCreateIntegrationContacts","Create integration contacts"],["capability","CanCreateIntegrationNotes","Create integration notes"],["capability (also relational)","CanEditPersonCampus","Edit person campus"],["capability","CanEditPersonDirectoryPrivacy","Edit person directory privacy"],["capability","CanEditPersonProfile","Edit person profile"],["capability","CanEditRegistrationSettings","Edit registration settings"],["capability","CanManageAllUserTokens","Manage all user tokens"],["capability","CanManageEmailQueue","Manage email queue"],["capability","CanManageNonContributions","Manage non contributions"],["capability","CanManageSystemNotifications","Manage system notifications"],["capability (also relational)","CanManageUserNotifications","Manage user notifications"],["capability","CanMergePeopleRecords","Merge people records"],["capability","CanRetrieveSpecialContent","Retrieve special content"],["capability","CanSendPushNotification","Send push notification"],["capability","CanStopOrDeleteNotification","Stop or delete notification"],["capability (also relational)","CanViewDirectoryPersonProfile","View directory person profile"]],"Privacy":[["role check","PrivacySettingsModel.cs","Privacy settings"]],"Public":[["role check","IncomingSmsModel.cs","Incoming SMS"],["role check","MobileAPIv2Controller.cs","Mobile AP iv2"],["role check","OrgContentInfo.cs","Org content info"],["screen element","ScriptResults.cshtml","Org content, Script results"]],"Rainforest (API)":[["API endpoint","v1/Rainforest/AddMerchant","Add merchant (rainforest)"],["API endpoint","v1/Rainforest/AdminSettings","Admin settings (rainforest)"],["API endpoint","v1/Rainforest/CreateMerchant","Create merchant (rainforest)"],["API endpoint","v1/Rainforest/DepositReportSession","Deposit report session (rainforest)"],["API endpoint","v1/Rainforest/MerchantOnboardingSession","Merchant onboarding session (rainforest)"],["API endpoint","v1/Rainforest/Merchants","Merchants (rainforest)"],["API endpoint","v1/Rainforest/OnboardingMerchants","Onboarding merchants (rainforest)"],["API endpoint","v1/Rainforest/PayinDetailsSession","Payin details session (rainforest)"]],"Registrations (API)":[["API endpoint","v1/Involvements/{orgId:int}/Members/{peopleId}/TransactionSummary","Transaction summary (involvements members), for a given org and people"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/Settings","Settings (involvements registrations), for a given org"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/Settings/Settings","Settings (involvements registrations settings), for a given org"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/{peopleId:int}/Form","Form (involvements registrations), for a given org and people"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/{peopleId:int}/Forms","Forms (involvements registrations), for a given org and people"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/{regId:guid}/Form","Form (involvements registrations), for a given org and reg"],["API endpoint","v1/Involvements/{orgId:int}/TransactionSummary","Transaction summary (involvements), for a given org"]],"Reports":[["role check","CustomReportsModel.cs","Custom reports"],["page or action","ReportsController.AddReport","Add report"],["page or action","ReportsController.Application","Application reports"],["page or action","ReportsController.DeleteCustomReport","Delete custom report"],["page or action","ReportsController.EditCustomReport","Edit custom report"],["page or action","ReportsController.RegistrationSummary","Registration summary"],["screen element","SqlReport.cshtml","Reports, SQL report"],["role check","StatusFlagsExportModel.cs","Status flags export"]],"Reservables":[["role check","ReservablesController.cs","Reservables"]],"Search":[["screen element","Edit.cshtml","Saved query, Edit"],["screen element","Index.cshtml","Saved query"],["role check","PictureDirectoryController.cs","Picture directory"],["role check","PictureDirectoryModel.cs","Picture directory"],["role check","RegistrationSearchModel.cs","Registration search"],["screen element","Row.cshtml","Saved query, Row"],["screen element","SaveAs.cshtml","Query, Save as"],["screen element","SelectCondition.cshtml","Query, Select condition"]],"Setup":[["page or action","BibleBookController","Bible book controller"],["page or action","DashboardWidgetController.Delete","Delete dashboard widget"],["page or action","DashboardWidgetController.Index","Index dashboard widget"],["page or action","DashboardWidgetController.Manage","Manage dashboard widget"],["page or action","DashboardWidgetController.Reorder","Reorder dashboard widget"],["page or action","DashboardWidgetController.Toggle","Toggle dashboard widget"],["page or action","DashboardWidgetController.Update","Update dashboard widget"],["page or action","DivisionController.Index","Index division"],["page or action","EmailDelegationController","Email delegation controller"],["screen element","Index.cshtml","Dashboard widget"],["page or action","KeywordExtraValuesController","Keyword extra values controller"],["page or action","KeywordsController","Keywords controller"],["screen element","List.cshtml","Lookup, List"],["page or action","LookupController","Lookup controller"],["role check","LookupController.cs","Lookup"],["page or action","MemberTypeController","Member type controller"],["page or action","MetroZipController","Metro zip controller"],["page or action","MinistryController","Ministry controller"],["page or action","ProgramController","Program controller"],["page or action","PromotionController","Promotion controller"],["page or action","RecurrenceRuleController.Index","Index recurrence rule"],["page or action","ResourceCategoryController","Resource category controller"],["page or action","ResourceMediaTypeController","Resource media type controller"],["page or action","ResourceTypeController","Resource type controller"],["page or action","RolesController","Roles controller"],["page or action","SettingController","Setting controller"],["page or action","TitleCodeController","Title code controller"],["page or action","TranslationsController.GetLanguages","Languages"],["page or action","TranslationsController.GetTranslations","Translations"],["page or action","TranslationsController.Index","Index translations"],["page or action","TranslationsController.RevertToDefault","Revert to default"],["page or action","TranslationsController.UpdateTranslation","Update translation"]],"Web":[["screen element","AdminAlerts.cshtml","Shared, Admin alerts"],["role check","AdminMenuVisibility.cs","Admin menu visibility"],["screen element","Custom.cshtml","Shared, Toolbar, Custom"],["screen element","EditTaskNote.vue","People, Person, Touchpoints, Edit task note"],["screen element","Entry.cshtml","Extra value, Entry"],["screen element","GearStandard.cshtml","Shared, Toolbar, Gear standard"],["page or action","HomeController.ActiveRecords","Active records"],["screen element","Location.cshtml","Extra value, Location"],["screen element","Notices.cshtml","Shared, Notices"],["screen element","People.cshtml","Shared, Menu, People"],["screen element","PyScript.cshtml","Script, Py script"],["screen element","RunPythonScriptProgress.cshtml","Script, Run python script progress"],["screen element","RunScript.cshtml","Script, Run script"],["screen element","Standard.cshtml","Extra value, Standard"],["screen element","Summary.cshtml","Extra value, Reports, Summary"],["screen element","Support.cshtml","Home, Support"],["role check","ToolbarController.cs","Toolbar"]]},"ApiOnly":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Manage":[["role check","AccountModel.cs","Account"]],"Other":[["role check","ContributionServiceDataWarehouseTests.cs","Contribution service data warehouse tests"]],"Permission checks":[["capability","CanRetrieveDataWarehouseList","Retrieve data warehouse list"]]},"ApiWrite":{"Web":[["page or action","PostController.cs.AddContribution","Add contribution"],["page or action","PostController.cs.ReverseContribution","Reverse contribution"]]},"AppAdmin":{"Core services":[["role check","ApplicationSettingsService.cs","Application settings"],["role check","PermissionsService.cs","Permissions"]],"Manage":[["screen element","Index.cshtml","Display"]],"Org":[["screen element","RegistrationFormRegistration.cshtml","Org, Registration form registration"],["screen element","RegistrationFormSettings.cshtml","Org, Registration form settings"]],"People":[["screen element","Tab.cshtml","Person, Communications"]],"Permission checks":[["capability","CanAddOrganizationPosts","Add organization posts"],["capability (also relational)","CanArchivePrayerRequest","Archive prayer request"],["capability (also relational)","CanChangePostCommentStatus","Change post comment status"],["capability (also relational)","CanChangeUserPostStatus","Change user post status"],["capability","CanDeleteOrganizationPosts","Delete organization posts"],["capability (also relational)","CanDeletePrayerRequest","Delete prayer request"],["capability","CanDeleteUserPosts","Delete user posts"],["capability","CanEditOrganizationPosts","Edit organization posts"],["capability (also relational)","CanModifyPrayerRequestAnswered","Modify prayer request answered"]],"Setup":[["page or action","SettingController","Setting controller"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"Attendance":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Org":[["screen element","Attendance.cshtml","Meeting, Attendance"],["screen element","Email.cshtml","Org search, Toolbar, Email"],["screen element","Gear.cshtml","Org search, Toolbar, Gear"],["screen element","Index.cshtml","Meeting"],["page or action","MeetingController.CreateMeeting","Create meeting"],["page or action","MeetingController.MarkAttendance","Mark attendance"],["page or action","MeetingController.MarkRegistered","Mark registered"],["page or action","MeetingController.ScanTicket","Scan ticket"],["role check","MeetingController.cs","Meeting"],["screen element","MeetingHeader.cshtml","Meeting, Meeting header"],["screen element","MeetingRightSideBar.cshtml","Meeting, Meeting right side bar"],["screen element","Meetings.cshtml","Org, Meetings"],["page or action","OrgSearchController.DisplayAttendanceNotices","Display attendance notices"],["page or action","OrgSearchController.EmailAttendanceNotices","Email attendance notices"],["role check","OrganizationModel.cs","Organization"],["role check","SettingsAttendanceModel.cs","Settings attendance"],["screen element","iPad.cshtml","Meeting, I pad"]],"Permission checks":[["capability (also relational)","CanTakeAttendance","Take attendance"],["capability (also relational)","CanTakeMeetingAttendance","Take meeting attendance"],["capability (also relational)","CanTakeOrganizationAttendance","Take organization attendance"]],"Public":[["role check","MobileAPIv2Controller.cs","Mobile AP iv2"]],"Reservables":[["role check","ReservablesController.cs","Reservables"]]},"BackgroundCheck":{"People":[["role check","EnrollmentController.cs","Enrollment"],["screen element","Index.cshtml","Volunteering"],["screen element","Volunteer.cshtml","Person, Enrollment, Volunteer"]],"Web":[["screen element","GearStandard.cshtml","Shared, Toolbar, Gear standard"],["role check","ToolbarController.cs","Toolbar"]]},"CalendarManagement":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Org":[["role check","MeetingController.cs","Meeting"],["role check","OrganizationModel.cs","Organization"],["role check","SettingsAttendanceModel.cs","Settings attendance"]],"Other":[["role check","NavBarPermissions.cs","Nav bar permissions"]],"Permission checks":[["capability","CanApproveMeetings","Approve meetings"],["capability","CanApproveReservable","Approve reservable"]],"Reservables":[["role check","ReservablesController.cs","Reservables"]],"Setup":[["screen element","List.cshtml","Lookup, List"]]},"Checkin":{"CheckIn":[["page or action","CheckInDashboardController.CheckInDashboard","Check in dashboard"],["page or action","CheckinSetupController","Checkin setup controller"],["page or action","ClassroomDashboardController.Index","Index classroom dashboard"],["page or action","ClassroomDashboardController.Person","Person classroom dashboard"]],"Involvements (API)":[["API endpoint","v1/Involvements","Involvements"],["API endpoint","v1/Involvements/Divisions/Search","Search (involvements divisions)"]],"Meetings (API)":[["API endpoint","v1/Involvements/{orgId:int}/Meetings","Meetings (involvements), for a given org"]],"Org":[["role check","OrgPeopleModel.cs","Org people"]],"People":[["role check","PictureResult.cs","Picture result"]],"Public":[["role check","CheckInAPIv2Controller.cs","Check in AP iv2"],["role check","MobileAPIv2Controller.cs","Mobile AP iv2"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"],["screen element","Organization.cshtml","Shared, Menu, Organization"]]},"Checkout":{"CheckIn":[["screen element","Index.cshtml","Classroom dashboard"]],"Public":[["role check","CheckInAPIv2Controller.cs","Check in AP iv2"]]},"ContentEdit":{"Public":[["role check","OrgContentInfo.cs","Org content info"]]},"Conversion":{"Org":[["role check","OrgPeopleModel.cs","Org people"]]},"Coupon":{"Main":[["page or action","CouponController","Coupon controller"]],"Manage":[["role check","TransactionsModel.cs","Transactions"]],"OnlineReg":[["role check","OnlineRegController.OnePageReg.cs","Online reg controller one page reg"],["role check","OnlineRegModel.Confirm.cs","Online reg model confirm"]],"Web":[["screen element","Organization.cshtml","Shared, Menu, Organization"]]},"Coupon2":{"Main":[["screen element","List.cshtml","Coupon, List"]]},"Delete":{"Org":[["screen element","Gear.cshtml","Org, Toolbar, Gear"],["page or action","OrgController.Delete","Delete org"]],"People":[["screen element","Gear.cshtml","Person, Toolbar, Gear"]]},"Design":{"Org":[["screen element","MessagesEdit.cshtml","Org, Registration, Messages edit"]],"Reports":[["screen element","SqlReport.cshtml","Reports, SQL report"]]},"Developer":{"ContentSafety (API)":[["API endpoint","v1/ContentSafety/CheckImageContent","Check image content (content safety)"],["API endpoint","v1/ContentSafety/CheckTextContent","Check text content (content safety)"]],"Core services":[["role check","ApplicationSettingsService.cs","Application settings"],["role check","PermissionsService.cs","Permissions"],["role check","TransactionsService.cs","Transactions"]],"Dialog":[["screen element","Display.cshtml","Org member dialog, Display"],["screen element","Edit.cshtml","Org member dialog, Edit"],["screen element","History.cshtml","Transaction history, History"],["screen element","Index.cshtml","Org members update"],["screen element","MemberData.cshtml","Org member dialog, Tabs, Member data"],["page or action","OrgMemberDialogController.DeleteQuestion","Delete question"],["screen element","Questions.cshtml","Org member dialog, Tabs, Questions"]],"Finance":[["screen element","Index.cshtml","Bundle"],["screen element","TotalsByFund.cshtml","Finance reports, Totals by fund"]],"Main":[["screen element","Announcement.cshtml","Announcement, Announcement"]],"Manage":[["role check","AccountController.cs","Account"],["role check","EmailsController.cs","Emails"],["screen element","Index.cshtml","Upload people"],["screen element","List.cshtml","Transactions, List"],["screen element","OtherDeveloperActions.cshtml","Batch, Other developer actions"],["page or action","TransactionsController.AssignGoer","Assign goer"],["page or action","UploadExcelIpsController","Upload excel ips controller"],["page or action","UploadPeopleController","Upload people controller"]],"Org":[["screen element","Gear.cshtml","Org, Toolbar, Gear"],["role check","OrgPeopleModel.cs","Org people"],["screen element","Reports.cshtml","Org search, Toolbar, Reports"],["screen element","Tabs.cshtml","Org search, Tabs"]],"Other":[["role check","ContributionServiceDataWarehouseTests.cs","Contribution service data warehouse tests"],["role check","GatewayUtils.cs","Gateway utils"],["role check","ReplacementCodeService_PersonalInfoTests.cs","Replacement code service personal info tests"],["role check","TasksNotesModel.cs","Tasks notes"]],"People":[["role check","FailedMailModel.cs","Failed mail"],["screen element","GearStandard.cshtml","Person, Toolbar, Gear standard"],["role check","PersonController.Admin.cs","Person controller admin"],["screen element","UserEdit.cshtml","Person, System, User edit"],["screen element","Users.cshtml","Person, System, Users"]],"Permission checks":[["capability","CanRetrieveDataWarehouseList","Retrieve data warehouse list"],["capability","CanSeeTestTransactions","See test transactions"]],"Reports":[["screen element","SqlReport.cshtml","Reports, SQL report"]],"Search":[["page or action","ContactSearchController.DeleteContactsForType","Delete contacts for type"],["screen element","Index.cshtml","Saved query"],["screen element","Row.cshtml","Saved query, Row"]],"Setup":[["role check","GatewayController.cs","Gateway"],["screen element","GroupDialog.cshtml","SMS management, Group dialog"],["screen element","Index.cshtml","Setting"],["screen element","_Settings.cshtml","Setting, Settings"],["screen element","_SettingsSearchResults.cshtml","Setting, Settings search results"],["screen element","_SettingsTabs.cshtml","Setting, Settings tabs"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"],["role check","CodeValueModel.cs","Code value"],["screen element","EditTaskNote.vue","People, Person, Touchpoints, Edit task note"],["screen element","GearStandard.cshtml","Shared, Toolbar, Gear standard"]]},"Edit":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Dialog":[["page or action","AddOrganizationController","Add organization controller"],["screen element","Display.cshtml","Org member dialog, Display"],["screen element","Entry.cshtml","Org member dialog, Tabs, Extra value, Entry"],["screen element","Groups.cshtml","Org member dialog, Tabs, Groups"],["screen element","MemberData.cshtml","Org prev member dialog, Tabs, Member data"],["role check","OrgMemberModel.cs","Org member"],["screen element","Questions.cshtml","Org member dialog, Tabs, Questions"]],"ESignature (API)":[["API endpoint","v1/ESignature/ConfigureWebhook","Configure webhook (e signature)"]],"Giving":[["page or action","GivingManagementController.GetOnlineNotifyPersonList","Online notify person list"]],"Involvements (API)":[["API endpoint","v1/Involvements","Involvements"],["API endpoint","v1/Involvements/Divisions/Search","Search (involvements divisions)"]],"Main":[["screen element","Announcement.cshtml","Announcement, Announcement"],["screen element","Compose.cshtml","Email, Compose"],["screen element","Index.cshtml","Email"],["screen element","Options.cshtml","SMS, Options"]],"Manage":[["screen element","ManageArea.cshtml","Volunteers, Manage area"],["screen element","OtherDeveloperActions.cshtml","Batch, Other developer actions"],["page or action","TransactionsController","Transactions controller"],["role check","TransactionsController.cs","Transactions"]],"MeetingCategory":[["page or action","MeetingCategoryController.Edit","Edit meeting category"]],"Meetings (API)":[["API endpoint","v1/Involvements/{orgId:int}/Meetings","Meetings (involvements), for a given org"]],"Org":[["screen element","Attendance.cshtml","Org, Settings, Attendance"],["screen element","CheckIn.cshtml","Org, Settings, Check in"],["screen element","Directory.cshtml","Org, Settings, Directory"],["screen element","Email.cshtml","Org, Toolbar, Email"],["screen element","ExtrasGrid.cshtml","Org, Settings, Extras grid"],["screen element","Fees.cshtml","Org, Registration, Fees"],["screen element","Gear.cshtml","Org, Toolbar, Gear"],["screen element","General.cshtml","Org, Settings, General"],["screen element","Index.cshtml","Org"],["screen element","InlineCode.cshtml","Org, Display templates, Inline code"],["screen element","MeetingDetails.cshtml","Meeting, Meeting details"],["screen element","MeetingRightSideBar.cshtml","Meeting, Meeting right side bar"],["screen element","Messages.cshtml","Org, Registration, Messages"],["screen element","MobileSettings.cshtml","Org, Registration, Mobile settings"],["page or action","OrgController.CopySettings","Copy settings"],["page or action","OrgController.DeleteExtra","Delete extra"],["page or action","OrgController.EditExtra","Edit extra"],["page or action","OrgController.FeesEdit","Fees edit"],["page or action","OrgController.MessagesEdit","Messages edit"],["page or action","OrgController.MobileSettingsEdit","Mobile settings edit"],["page or action","OrgController.NewExtraValue","New extra value"],["page or action","OrgController.QuestionsEdit","Questions edit"],["page or action","OrgController.RegistrationEdit","Registration edit"],["role check","OrgPeopleModel.cs","Org people"],["role check","OrgSearchController.cs","Org search"],["screen element","Questions.cshtml","Org, Registration, Questions"],["page or action","RegSettingController.Update","Update reg setting"],["screen element","Registration.cshtml","Org, Registration, Registration"],["screen element","RegistrationFormRegistration.cshtml","Org, Registration form registration"],["screen element","Results.cshtml","Org search, Results"],["screen element","SevenDayView.cshtml","Org, Volunteer scheduler, Seven day view"],["role check","VolunteerSchedulerModel.cs","Volunteer scheduler"]],"People":[["screen element","Display.cshtml","Person, Profile, Comments, Display"],["screen element","Gear.cshtml","Person, Toolbar, Gear"],["screen element","GearStandard.cshtml","Person, Toolbar, Gear standard"],["screen element","Index.cshtml","Attributes"],["screen element","InlineCode.cshtml","Person, Display templates, Inline code"],["screen element","Optouts.cshtml","Person, Communications, Optouts"],["screen element","Registrations.cshtml","Person, Enrollment, Registrations"],["screen element","RegistrationsEdit.cshtml","Person, Enrollment, Registrations edit"]],"Permission checks":[["capability (also relational)","CanCheckInFamily","Check in family"],["capability","CanCreateIntegrationContacts","Create integration contacts"],["capability","CanCreateIntegrationNotes","Create integration notes"],["capability","CanEditPersonDirectoryPrivacy","Edit person directory privacy"],["capability","CanEditPersonProfile","Edit person profile"],["capability","CanEditRegistrationSettings","Edit registration settings"],["capability (also relational)","CanViewDirectoryPersonProfile","View directory person profile"]],"Public":[["role check","OrgContentInfo.cs","Org content info"]],"Registrations (API)":[["API endpoint","v1/Involvements/{orgId:int}/Members/{peopleId}/TransactionSummary","Transaction summary (involvements members), for a given org and people"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/Settings","Settings (involvements registrations), for a given org"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/Settings/Settings","Settings (involvements registrations settings), for a given org"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/{peopleId:int}/Form","Form (involvements registrations), for a given org and people"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/{peopleId:int}/Forms","Forms (involvements registrations), for a given org and people"],["API endpoint","v1/Involvements/{orgId:int}/Registrations/{regId:guid}/Form","Form (involvements registrations), for a given org and reg"],["API endpoint","v1/Involvements/{orgId:int}/TransactionSummary","Transaction summary (involvements), for a given org"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"],["screen element","Entry.cshtml","Extra value, Entry"],["screen element","GearStandard.cshtml","Shared, Toolbar, Gear standard"],["screen element","NavBar.cshtml","Shared, Nav bar"],["screen element","Organization.cshtml","Shared, Menu, Organization"],["role check","ToolbarController.cs","Toolbar"],["screen element","User.cshtml","Shared, Menu, User"]]},"EditCampus":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Permission checks":[["capability (also relational)","CanEditPersonCampus","Edit person campus"]]},"EmailTemplates":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Manage":[["page or action","EmailTemplatesController","Email templates controller"],["screen element","Index.cshtml","Email templates"]],"Permission checks":[["capability","CanRetrieveSpecialContent","Retrieve special content"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"EmailTest":{"Main":[["role check","MassEmailer.cs","Mass emailer"]],"Manage":[["role check","EmailsController.cs","Emails"]]},"Finance":{"Core services":[["role check","PermissionsService.cs","Permissions"],["role check","TransactionsService.cs","Transactions"]],"Dialog":[["screen element","Edit.cshtml","Org member dialog, Edit"],["screen element","GetCheckImage.cshtml","Check image, Check image"],["role check","GetCheckImageController.cs","Check image"],["screen element","MemberData.cshtml","Org member dialog, Tabs, Member data"]],"Finance":[["page or action","BundleController","Bundle controller"],["page or action","FinanceReportsController","Finance reports controller"],["page or action","FinanceReportsController.DeleteScheduledGift","Delete scheduled gift"],["page or action","FinanceReportsController.ManagedGiving","Managed giving"],["role check","FinanceReportsController.cs","Finance reports"],["page or action","FundController","Fund controller"],["page or action","FundController.Index","Index fund"],["page or action","FundSetsController","Fund sets controller"],["screen element","ManagedGiving.cshtml","Finance reports, Managed giving"],["page or action","PostBundleController","Post bundle controller"],["page or action","QuickBooksController","Quick books controller"],["screen element","TotalsByFund.cshtml","Finance reports, Totals by fund"],["role check","TotalsByFundModel.cs","Totals by fund"],["role check","Transactions2Controller.cs","Transactions2"]],"Finance (API)":[["API endpoint","v2/FinanceReports/LapseGiving","Lapse giving (finance reports)"]],"Giving":[["page or action","GivingManagementController.CheckUrlAvailability","Check URL availability"],["page or action","GivingManagementController.Create","Create giving management"],["page or action","GivingManagementController.Delete","Delete giving management"],["page or action","GivingManagementController.GetAvailableFunds","Available funds"],["page or action","GivingManagementController.GetCampusList","Campus list"],["page or action","GivingManagementController.GetConfirmationEmailList","Confirmation email list"],["page or action","GivingManagementController.GetEntryPoints","Entry points"],["page or action","GivingManagementController.GetOnlineNotifyPersonList","Online notify person list"],["page or action","GivingManagementController.GetShellList","Shell list"],["page or action","GivingManagementController.Index","Index giving management"],["page or action","GivingManagementController.List","List giving management"],["page or action","GivingManagementController.Manage","Manage giving management"],["page or action","GivingManagementController.New","New giving management"],["page or action","GivingManagementController.SaveGivingPageEnabled","Save giving page enabled"],["page or action","GivingManagementController.SetGivingDefaultPage","Set giving default page"],["page or action","GivingManagementController.Update","Update giving management"],["role check","GivingPaymentModel.cs","Giving payment"]],"Giving (API)":[["API endpoint","v1/Giving/ContributionFunds","Contribution funds (giving)"]],"Main":[["role check","AdminController.cs","Admin"]],"Manage":[["role check","AccountModel.cs","Account"],["page or action","BatchController.RetrieveBatchData","Retrieve batch data"],["screen element","Details.cshtml","Emails, Details"],["role check","EmailModel.cs","Email"],["role check","EmailsController.cs","Emails"],["role check","EmailsModel.cs","Emails"],["page or action","TransactionsController.Adjust","Adjust transactions"],["page or action","TransactionsController.AssignGoer","Assign goer"],["page or action","TransactionsController.CreditVoid","Credit void"],["page or action","TransactionsController.CreditVoidAjax","Credit void ajax"],["page or action","TransactionsController.DeleteGoerSenderAmount","Delete goer sender amount"],["page or action","TransactionsController.DeleteManual","Delete manual"],["role check","TransactionsModel.cs","Transactions"]],"Org":[["screen element","Fees.cshtml","Org, Registration, Fees"],["screen element","FeesEdit.cshtml","Org, Registration, Fees edit"],["screen element","RegistrationFormSettings.cshtml","Org, Registration form settings"]],"Other":[["role check","ContributionFundExtensions.cs","Contribution fund"],["role check","ContributionServiceDataWarehouseTests.cs","Contribution service data warehouse tests"],["role check","DataStoreService_Contributions.cs","Data store service contributions"],["role check","DbUtil.Context.cs","Db util context"],["role check","GivingConfirmationBuilder.cs","Giving confirmation builder"],["role check","GivingToFund.cs","Giving to fund"],["role check","Person.cs","Person"],["role check","ReplacementCodeService_GivingFundTests.cs","Replacement code service giving fund tests"],["role check","ScheduledGiving.cs","Scheduled giving"],["role check","TouchPointAuthenticationMiddleware.cs","Touch point authentication middleware"]],"People":[["screen element","CombineGiving.cshtml","Shared, Combine giving"],["screen element","Contributions.cshtml","Person, Giving, Contributions"],["role check","ContributionsModel.cs","Contributions"],["role check","EmailModel.cs","Email"],["page or action","PersonController.DeleteDocument","Delete document"],["page or action","PersonController.DeletePledge","Delete pledge"],["page or action","PersonController.EditPledge","Edit pledge"],["page or action","PersonController.FinanceDocuments","Finance documents"],["page or action","PersonController.MemberDocumentUpdateName","Member document update name"],["page or action","PersonController.MergePledge","Merge pledge"],["page or action","PersonController.UploadDocument","Upload document"],["role check","PersonController.cs","Person"],["role check","PictureResult.cs","Picture result"],["role check","ProfileController.cs","Profile"],["screen element","Statements.cshtml","Person, Giving, Statements"],["screen element","StatementsWithFund.cshtml","Person, Giving, Statements with fund"],["role check","SystemController.cs","System"],["screen element","Tab.cshtml","Person, Giving"],["screen element","UserEdit.cshtml","Person, System, User edit"]],"Permission checks":[["capability","CanManageContributions","Manage contributions"],["capability","CanManageNonContributions","Manage non contributions"],["capability","CanManageOrViewRecurringOnlineGiving","Manage or view recurring online giving"],["capability","CanRetrieveContributionFunds","Retrieve contribution funds"],["capability","CanRetrieveDataWarehouseList","Retrieve data warehouse list"],["capability","CanUserRunFinanceReports","User run finance reports"]],"Public":[["role check","CheckScanAPIController.cs","Check scan API"],["role check","CheckScanAuthentication.cs","Check scan authentication"],["role check","MobileAuthentication.cs","Mobile authentication"]],"Rainforest (API)":[["API endpoint","v1/Rainforest/CleanupMigrationLeftovers","Cleanup migration leftovers (rainforest)"],["API endpoint","v1/Rainforest/DepositReportSession","Deposit report session (rainforest)"],["API endpoint","v1/Rainforest/MerchantOnboardingSession","Merchant onboarding session (rainforest)"],["API endpoint","v1/Rainforest/Merchants","Merchants (rainforest)"],["API endpoint","v1/Rainforest/OnboardingMerchants","Onboarding merchants (rainforest)"],["API endpoint","v1/Rainforest/PayinDetailsSession","Payin details session (rainforest)"],["API endpoint","v1/Rainforest/ProcessMigrationEvents","Process migration events (rainforest)"],["API endpoint","v1/Rainforest/ProcessWebhookEvents","Process webhook events (rainforest)"],["API endpoint","v1/Rainforest/Refund","Refund (rainforest)"]],"Reports":[["page or action","ExportController.Contributions","Contributions export"]],"Search":[["role check","PeopleSearchModel.cs","People search"],["role check","SearchController.cs","Search"]],"Setup":[["page or action","LookupController","Lookup controller"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"],["page or action","Home2Controller.TurnFinanceOff","Turn finance off"],["page or action","Home2Controller.TurnFinanceOn","Turn finance on"],["screen element","User.cshtml","Shared, Menu, User"]]},"FinanceAdmin":{"Core services":[["role check","PermissionsService.cs","Permissions"],["role check","TransactionsService.cs","Transactions"]],"Dialog":[["screen element","Edit.cshtml","Org member dialog, Edit"],["screen element","MemberData.cshtml","Org member dialog, Tabs, Member data"]],"Finance":[["role check","BatchesViewModel.cs","Batches view"],["role check","BundleModel.cs","Bundle"],["page or action","FinanceReportsController.DeleteScheduledGift","Delete scheduled gift"],["page or action","FinanceReportsController.ManagedGiving","Managed giving"],["role check","FinanceReportsController.cs","Finance reports"],["screen element","ManagedGiving.cshtml","Finance reports, Managed giving"],["screen element","TotalsByFund.cshtml","Finance reports, Totals by fund"]],"Finance (API)":[["API endpoint","v2/FinanceReports/LapseGiving","Lapse giving (finance reports)"]],"Giving (API)":[["API endpoint","v1/Giving/ContributionFunds","Contribution funds (giving)"]],"Manage":[["page or action","BatchController.RetrieveBatchData","Retrieve batch data"],["role check","EmailsModel.cs","Emails"]],"Org":[["screen element","Fees.cshtml","Org, Registration, Fees"],["screen element","FeesEdit.cshtml","Org, Registration, Fees edit"]],"Other":[["role check","ContributionServiceDataWarehouseTests.cs","Contribution service data warehouse tests"]],"People":[["role check","ContributionsModel.cs","Contributions"],["page or action","PersonController.DeleteDocument","Delete document"],["page or action","PersonController.FinanceDocuments","Finance documents"],["page or action","PersonController.MemberDocumentUpdateName","Member document update name"],["page or action","PersonController.UploadDocument","Upload document"],["role check","PictureResult.cs","Picture result"],["role check","ProfileController.cs","Profile"]],"Permission checks":[["capability","CanRetrieveContributionFunds","Retrieve contribution funds"],["capability","CanRetrieveDataWarehouseList","Retrieve data warehouse list"],["capability","CanUserRunFinanceReports","User run finance reports"]],"Rainforest (API)":[["API endpoint","v1/Rainforest/CleanupMigrationLeftovers","Cleanup migration leftovers (rainforest)"],["API endpoint","v1/Rainforest/ProcessMigrationEvents","Process migration events (rainforest)"],["API endpoint","v1/Rainforest/ProcessWebhookEvents","Process webhook events (rainforest)"],["API endpoint","v1/Rainforest/Refund","Refund (rainforest)"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"FinanceDataEntry":{"Dialog":[["screen element","GetCheckImage.cshtml","Check image, Check image"],["role check","GetCheckImageController.cs","Check image"]],"Finance":[["role check","BatchesController.cs","Batches"],["role check","BatchesModel.cs","Batches"],["role check","BatchesViewModel.cs","Batches view"],["page or action","BundleController","Bundle controller"],["role check","BundleController.cs","Bundle"],["role check","BundleModel.cs","Bundle"],["role check","BundlesModel.cs","Bundles"],["screen element","Display.cshtml","Bundle, Display"],["screen element","Index.cshtml","Bundle"],["role check","PledgesViewModel.cs","Pledges view"],["page or action","PostBundleController","Post bundle controller"],["role check","PostBundleController.cs","Post bundle"]],"Public":[["role check","CheckScanAPIController.cs","Check scan API"]],"Rainforest (API)":[["API endpoint","v1/Rainforest/DepositReportSession","Deposit report session (rainforest)"],["API endpoint","v1/Rainforest/Merchants","Merchants (rainforest)"],["API endpoint","v1/Rainforest/PayinDetailsSession","Payin details session (rainforest)"]],"Search":[["role check","PeopleSearchModel.cs","People search"],["role check","SearchController.cs","Search"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"],["role check","CodeValueModel.cs","Code value"]]},"FinanceViewOnlyDetail":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Finance":[["page or action","FinanceReportsController","Finance reports controller"],["page or action","FundController","Fund controller"],["screen element","TotalsByFund.cshtml","Finance reports, Totals by fund"]],"Giving":[["page or action","GivingManagementController.CheckUrlAvailability","Check URL availability"],["page or action","GivingManagementController.Create","Create giving management"],["page or action","GivingManagementController.Delete","Delete giving management"],["page or action","GivingManagementController.GetAvailableFunds","Available funds"],["page or action","GivingManagementController.GetCampusList","Campus list"],["page or action","GivingManagementController.GetConfirmationEmailList","Confirmation email list"],["page or action","GivingManagementController.GetEntryPoints","Entry points"],["page or action","GivingManagementController.GetOnlineNotifyPersonList","Online notify person list"],["page or action","GivingManagementController.GetShellList","Shell list"],["page or action","GivingManagementController.Index","Index giving management"],["page or action","GivingManagementController.List","List giving management"],["page or action","GivingManagementController.Manage","Manage giving management"],["page or action","GivingManagementController.New","New giving management"],["page or action","GivingManagementController.SaveGivingPageEnabled","Save giving page enabled"],["page or action","GivingManagementController.SetGivingDefaultPage","Set giving default page"],["page or action","GivingManagementController.Update","Update giving management"]],"Other":[["role check","ContributionFundExtensions.cs","Contribution fund"],["role check","Person.cs","Person"]],"People":[["role check","ContributionsModel.cs","Contributions"]],"Permission checks":[["capability","CanViewContributions","View contributions"]],"Reports":[["page or action","ExportController.Contributions","Contributions export"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"FundManager":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Finance":[["screen element","TotalsByFund.cshtml","Finance reports, Totals by fund"],["role check","TotalsByFundModel.cs","Totals by fund"]],"Other":[["role check","ContributionFundExtensions.cs","Contribution fund"],["role check","Person.cs","Person"]],"People":[["role check","ContributionsModel.cs","Contributions"]],"Permission checks":[["capability","CanManageOrViewLimitedFunds","Manage or view limited funds"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"GivingEmailTemplates":{"Giving":[["role check","GivingManagementController.cs","Giving management"]]},"ManageApplication":{"People":[["screen element","Display.cshtml","Volunteering, Display"],["screen element","DisplayApproval.cshtml","Volunteering, Display approval"],["screen element","Index.cshtml","Volunteering"],["role check","PictureResult.cs","Picture result"],["screen element","Volunteer.cshtml","Person, Ministry, Volunteer"],["page or action","VolunteeringController.Delete","Delete volunteering"]]},"ManageBundles":{"Finance":[["page or action","BundleController.Index","Index bundle"],["page or action","BundlesController","Bundles controller"],["page or action","PostBundleController.FundTotals","Fund totals"],["page or action","PostBundleController.Index","Index post bundle"]]},"ManageChat":{"Conversations (API)":[["API endpoint","v1/Conversations/Managed","Managed (conversations)"],["API endpoint","v1/Conversations/Managed/{conversationId:int}/BlockedMessages","Blocked messages (conversations managed), for a given conversation"],["API endpoint","v1/Conversations/Managed/{conversationId:int}/Members","Members (conversations managed), for a given conversation"],["API endpoint","v1/Conversations/Managed/{conversationId:int}/Members/{memberId:int}/Add","Add (conversations managed members), for a given conversation and member"],["API endpoint","v1/Conversations/Managed/{conversationId:int}/Members/{memberId:int}/Block","Block (conversations managed members), for a given conversation and member"],["API endpoint","v1/Conversations/Managed/{conversationId:int}/Members/{memberId:int}/Delete","Delete (conversations managed members), for a given conversation and member"],["API endpoint","v1/Conversations/Managed/{conversationId:int}/Members/{memberId:int}/Unblock","Unblock (conversations managed members), for a given conversation and member"],["API endpoint","v1/Conversations/Managed/{conversationId:int}/Messages","Messages (conversations managed), for a given conversation"],["API endpoint","v1/Conversations/Managed/{conversationId:int}/ModerationItems","Moderation items (conversations managed), for a given conversation"],["API endpoint","v1/Conversations/Managed/{conversationId:int}/ReportedMessages","Reported messages (conversations managed), for a given conversation"]],"Core services":[["role check","RealTimeService.cs","Real time"]],"Other":[["role check","RealTimeServiceTests.cs","Real time service tests"]]},"ManageEmails":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Manage":[["screen element","Details.cshtml","Emails, Details"],["role check","EmailModel.cs","Email"],["role check","EmailsController.cs","Emails"],["role check","EmailsModel.cs","Emails"]],"People":[["role check","CommunicationsController.cs","Communications"]],"Permission checks":[["capability","CanManageEmailQueue","Manage email queue"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"ManageEvents":{"Org":[["screen element","Results.cshtml","Org search, Results"]]},"ManageGroups":{"Dialog":[["screen element","Groups.cshtml","Org member dialog, Tabs, Groups"],["screen element","Index.cshtml","Org members update"],["page or action","OrgMemberDialogController.SmallGroupChecked","Small group checked"],["screen element","Questions.cshtml","Org member dialog, Tabs, Questions"]],"Org":[["screen element","Gear.cshtml","Org, Toolbar, Gear"]]},"ManageOrgMembers":{"Manage":[["page or action","OrgMembersController","Org members controller"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"ManagePrivacy":{"Dialog":[["screen element","Display.cshtml","Org member dialog, Display"]],"Org":[["screen element","RegistrationFormRegistration.cshtml","Org, Registration form registration"],["screen element","Settings.cshtml","Org, Settings"]],"Other":[["role check","Person.cs","Person"]],"Privacy":[["role check","PrivacySettingsModel.cs","Privacy settings"]]},"ManageProcesses":{"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"],["screen element","People.cshtml","Shared, Menu, People"]]},"ManageResources":{"Manage":[["page or action","MediaController.CreateResource","Create resource"],["page or action","MediaController.CreateResourceCategory","Create resource category"],["page or action","MediaController.CreateResourceType","Create resource type"],["page or action","MediaController.DeleteAttachment","Delete attachment"],["page or action","MediaController.DeleteResource","Delete resource"],["page or action","MediaController.EditResource","Edit resource"],["page or action","MediaController.MoveResource","Move resource"],["page or action","MediaController.MoveResourceCategoryList","Move resource category list"],["page or action","MediaController.SaveAttachmentListOrder","Save attachment list order"],["page or action","MediaController.SaveResourceListOrder","Save resource list order"],["page or action","MediaController.UpdateAttachment","Update attachment"],["page or action","MediaController.UploadAttachment","Upload attachment"],["page or action","ResourceController","Resource controller"]],"Setup":[["page or action","ResourceCategoryController","Resource category controller"],["page or action","ResourceMediaTypeController","Resource media type controller"],["page or action","ResourceTypeController","Resource type controller"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"ManageSMS":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Main":[["screen element","Announcement.cshtml","Announcement, Announcement"]],"Manage":[["screen element","Index.cshtml","SMS"]],"People":[["screen element","MessagesNotificationsLog.cshtml","Person, Communications, Messages notifications log"]],"Permission checks":[["capability","CanSendPushNotification","Send push notification"],["capability","CanStopOrDeleteNotification","Stop or delete notification"]],"Setup":[["screen element","Index.cshtml","SMS management"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"ManageTouchpoints":{"Other":[["role check","TasksNotesModel.cs","Tasks notes"]],"Web":[["screen element","EditTaskNote.vue","People, Person, Touchpoints, Edit task note"]]},"ManageTransactions":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Dialog":[["screen element","Edit.cshtml","Org member dialog, Edit"],["screen element","Index.cshtml","Org members update"],["screen element","MemberData.cshtml","Org member dialog, Tabs, Member data"],["screen element","Tickets.cshtml","Org member dialog, Tabs, Tickets"]],"Manage":[["page or action","BatchController.RetrieveBatchData","Retrieve batch data"],["page or action","TransactionsController","Transactions controller"],["page or action","TransactionsController.Adjust","Adjust transactions"],["page or action","TransactionsController.AssignGoer","Assign goer"],["page or action","TransactionsController.CreditVoid","Credit void"],["page or action","TransactionsController.CreditVoidAjax","Credit void ajax"],["page or action","TransactionsController.DeleteGoerSenderAmount","Delete goer sender amount"],["page or action","TransactionsController.DeleteManual","Delete manual"],["role check","TransactionsController.cs","Transactions"],["role check","TransactionsModel.cs","Transactions"]],"Permission checks":[["capability","CanManageNonContributions","Manage non contributions"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"Manager":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Manage":[["page or action","MergeController","Merge controller"]],"People":[["screen element","Display.cshtml","Person, Personal, Display"],["screen element","Gear.cshtml","Person, Toolbar, Gear"]],"Permission checks":[["capability","CanMergePeopleRecords","Merge people records"]]},"Manager2":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Manage":[["page or action","DuplicatesController","Duplicates controller"],["screen element","Index.cshtml","Merge"],["page or action","MergeController","Merge controller"]],"People":[["screen element","Gear.cshtml","Person, Toolbar, Gear"]],"Permission checks":[["capability","CanMergePeopleRecords","Merge people records"]]},"McpAccess":{"Manage":[["role check","McpAdminModel.cs","Mcp admin"]],"Other":[["role check","McpAuthenticator.cs","Mcp authenticator"]]},"McpViewContact":{"Manage":[["role check","McpAdminModel.cs","Mcp admin"]],"Other":[["role check","RedactionService.cs","Redaction"]]},"McpViewDemographics":{"Manage":[["role check","McpAdminModel.cs","Mcp admin"]],"Other":[["role check","RedactionService.cs","Redaction"]]},"MemberDocs":{"People":[["screen element","Display.cshtml","Person, Profile, Membership, Display"],["page or action","PersonController.DeleteDocument","Delete document"],["page or action","PersonController.MemberDocumentUpdateName","Member document update name"],["page or action","PersonController.MemberDocuments","Member documents"],["page or action","PersonController.UploadDocument","Upload document"],["role check","PictureResult.cs","Picture result"],["screen element","Tab.cshtml","Person, Profile"]]},"Membership":{"People":[["screen element","Display.cshtml","Person, Profile, Membership, Display"],["screen element","Edit.cshtml","Person, Personal, Edit"],["page or action","PersonController.DeleteDocument","Delete document"],["page or action","PersonController.MemberDocumentUpdateName","Member document update name"],["page or action","PersonController.MemberDocuments","Member documents"],["page or action","PersonController.UploadDocument","Upload document"],["role check","PictureResult.cs","Picture result"],["role check","ProfileController.cs","Profile"],["screen element","Tab.cshtml","Person, Profile"]]},"MissionGiving":{"Dialog":[["screen element","Edit.cshtml","Org member dialog, Edit"],["screen element","Index.cshtml","Org members update"],["screen element","MemberData.cshtml","Org member dialog, Tabs, Member data"]],"Org":[["screen element","Gear.cshtml","Org, Toolbar, Gear"],["screen element","Reports.cshtml","Org search, Toolbar, Reports"]]},"OrgLeadersOnly":{"Core services":[["role check","DirectoryService.cs","Directory"],["role check","OrganizationsService.cs","Organizations"],["role check","PermissionsService.cs","Permissions"],["role check","TaskNoteService.cs","Task note"]],"Dialog":[["role check","OrgMemberModel.cs","Org member"]],"Org":[["role check","OrganizationModel.cs","Organization"],["role check","VolunteerSchedulerModel.cs","Volunteer scheduler"]],"Other":[["role check","DataStoreService_Organizations.cs","Data store service organizations"],["role check","PeopleSearch.cs","People search"],["role check","Person.cs","Person"],["role check","TasksNotesModel.cs","Tasks notes"]],"People":[["screen element","Current.cshtml","Person, Enrollment, Current"],["screen element","Emails.cshtml","Person, Communications, Emails"],["screen element","Members.cshtml","Person, Family, Members"],["screen element","Pending.cshtml","Person, Enrollment, Pending"],["screen element","Previous.cshtml","Person, Enrollment, Previous"],["screen element","Related.cshtml","Person, Family, Related"],["screen element","Tab.cshtml","Person, Enrollment"]],"Permission checks":[["capability (also relational)","CanAddPersonTaskNotes","Add person task notes"],["capability (also relational)","CanTakeAttendance","Take attendance"],["capability (also relational)","CanTakeMeetingAttendance","Take meeting attendance"],["capability (also relational)","CanTakeOrganizationAttendance","Take organization attendance"],["capability (also relational)","CanViewPerson","View person"],["capability (also relational)","CanViewPersonBadges","View person badges"],["capability (also relational)","CanViewPersonChannels","View person channels"],["capability (also relational)","CanViewPersonEmergencyContact","View person emergency contact"],["capability (also relational)","CanViewPersonEngagementScore","View person engagement score"],["capability (also relational)","CanViewPersonInvolvements","View person involvements"],["capability (also relational)","CanViewPersonPrayerRequests","View person prayer requests"],["capability (also relational)","CanViewPersonTaskNotes","View person task notes"]],"Privacy":[["role check","PrivacySettingsModel.cs","Privacy settings"]],"Public":[["role check","MobileAPIv2Controller.cs","Mobile AP iv2"],["role check","MobileAuthentication.cs","Mobile authentication"]]},"OrgTagger":{"Org":[["screen element","Tabs.cshtml","Org search, Tabs"]]},"ScheduleEmails":{"Main":[["screen element","Announcement.cshtml","Announcement, Announcement"],["screen element","Compose.cshtml","Email, Compose"],["screen element","Index.cshtml","Email"],["screen element","Options.cshtml","SMS, Options"]]},"SchedulerTemplates":{"OnlineReg":[["role check","VolSubModel.cs","Vol sub"]],"Org":[["role check","VolunteerSchedulerModel.cs","Volunteer scheduler"],["role check","VolunteerSchedulerRequestModel.cs","Volunteer scheduler request"]]},"SendSMS":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Main":[["screen element","Announcement.cshtml","Announcement, Announcement"]],"Manage":[["screen element","Index.cshtml","SMS"]],"Org":[["screen element","Email.cshtml","Org, Toolbar, Email"]],"Other":[["role check","SmsHelper.cs","SMS helper"]],"People":[["screen element","MessagesNotificationsLog.cshtml","Person, Communications, Messages notifications log"]],"Permission checks":[["capability","CanSendPushNotification","Send push notification"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"],["screen element","Email.cshtml","Shared, Toolbar, Email"]]},"SpecialContentBasic":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Manage":[["page or action","InvolvementController","Involvement controller"]],"Permission checks":[["capability","CanRetrieveSpecialContent","Retrieve special content"]],"Reports":[["role check","CustomReportsModel.cs","Custom reports"],["page or action","ReportsController.AddReport","Add report"],["page or action","ReportsController.DeleteCustomReport","Delete custom report"],["page or action","ReportsController.EditCustomReport","Edit custom report"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"],["screen element","Custom.cshtml","Shared, Toolbar, Custom"],["role check","ToolbarController.cs","Toolbar"]]},"SpecialContentFull":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Manage":[["role check","DisplayController.cs","Display"],["screen element","EditPythonScript.cshtml","Display, Edit python script"],["screen element","EditSqlScript.cshtml","Display, Edit SQL script"],["screen element","Index.cshtml","Display"]],"Permission checks":[["capability","CanRetrieveSpecialContent","Retrieve special content"]],"Reports":[["role check","CustomReportsModel.cs","Custom reports"],["page or action","ReportsController.AddReport","Add report"],["page or action","ReportsController.DeleteCustomReport","Delete custom report"],["page or action","ReportsController.EditCustomReport","Edit custom report"]],"Setup":[["screen element","Index.cshtml","Dashboard widget"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"],["screen element","Custom.cshtml","Shared, Toolbar, Custom"],["role check","ToolbarController.cs","Toolbar"]]},"StatusFlag":{"Other":[["role check","Condition.Miscellaneous.cs","Condition miscellaneous"]],"Reports":[["role check","StatusFlagsExportModel.cs","Status flags export"]],"Search":[["screen element","Edit.cshtml","Saved query, Edit"],["screen element","Index.cshtml","Saved query"],["screen element","Row.cshtml","Saved query, Row"],["screen element","SaveAs.cshtml","Query, Save as"]]},"Support":{"Manage":[["role check","AccountModel.cs","Account"]],"People":[["screen element","UserEdit.cshtml","Person, System, User edit"],["screen element","Users.cshtml","Person, System, Users"]],"Web":[["screen element","AdminAlerts.cshtml","Shared, Admin alerts"],["role check","HomeController.cs","Home"],["screen element","Support.cshtml","Home, Support"]]},"TicketScanning":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Permission checks":[["capability","CanScanTickets","Scan tickets"]]},"Ticketing":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Dialog":[["screen element","Tickets.cshtml","Org member dialog, Tabs, Tickets"]],"Main":[["role check","CouponController.cs","Coupon"]],"OnlineReg":[["screen element","RegistrationForm.cshtml","Online reg, Registration form"]],"Org":[["screen element","Gear.cshtml","Org, Toolbar, Gear"],["screen element","GeneralEdit.cshtml","Org, Settings, General edit"],["screen element","RegistrationFormSettings.cshtml","Org, Registration form settings"]],"People (API)":[["API endpoint","v1/People/Search/AutocompleteEx/Registration","Registration (people search autocomplete ex)"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"TrainingClasses":{"People":[["role check","EnrollmentController.cs","Enrollment"],["screen element","Volunteer.cshtml","Person, Enrollment, Volunteer"]],"Web":[["screen element","GearStandard.cshtml","Shared, Toolbar, Gear standard"],["role check","ToolbarController.cs","Toolbar"]]},"UploadPeople":{"Manage":[["page or action","UploadPeopleController","Upload people controller"]]},"ViewApplication":{"People":[["screen element","Display.cshtml","Volunteering, Display"],["screen element","DisplayApproval.cshtml","Volunteering, Display approval"],["role check","PictureResult.cs","Picture result"],["screen element","Volunteer.cshtml","Person, Ministry, Volunteer"]]},"ViewPrivateTouchpoints":{"Other":[["role check","TasksNotesModel.cs","Tasks notes"]]},"ViewResources":{"People":[["screen element","Index.cshtml","Person"]]},"ViewTransactions":{"Core services":[["role check","PermissionsService.cs","Permissions"]],"Permission checks":[["capability","CanViewNonContributions","View non contributions"]],"Web":[["role check","AdminMenuVisibility.cs","Admin menu visibility"]]},"VolDocs":{"People":[["screen element","Volunteer.cshtml","Person, Enrollment, Volunteer"]]}}"""


# ===========================================================================
if handle_ajax():
    pass
else:
    try:
        if not current_user_id():
            model.Form = "<p>Sign in to use this.</p>"
        else:
            render()
    except Exception:
        model.Form = ("<h3>Access Audit failed to load</h3><pre>"
                      + esc(traceback.format_exc()) + "</pre>")
