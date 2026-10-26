#roles=Admin
# -*- coding: utf-8 -*-
"""
Leader Retention
================
How well are we keeping our leaders from one fiscal year to the next, in
each program and each primary division, and who should someone call?

For every program it shows, year over year:
  KEPT             led in the same program (or division) again next year
  STILL LEADING    led anywhere next year, including a different program
  MOVED            led somewhere else next year, but not here
  STOPPED          did not lead anywhere next year

Click any program to expand its primary divisions. Click any Stopped or
Moved number for the follow-up list: who they are, what they last led and
when, whether they still come to church, which leader rosters still list
them, and how to reach them.

WHO COUNTS AS A LEADER
----------------------
Anyone in one of the leader roles picked in Settings (by default every role
TouchPoint treats as a leader for attendance, except Staff Member).

Two ways to decide someone led in a year, picked in Settings:
  * Attended (default): they were marked present in a leader role at least
    once in that program during the year. Best measure, because leaders are
    rarely taken off rosters, so a roster says someone is still leading long
    after they stopped.
  * On roster: they were listed in a leader role for at least 30 days of
    the year. Use it for programs that do not take attendance.

Programs whose leaders serve for a week (VBS, camps) pull the church-wide
rate down because most of those volunteers only serve that one week. Leave
them out with the Settings exclusions.

Everything here is read-only. Settings are saved to the Special Content
text record "TPxi_LeaderRetention_Settings".

--Upload Instructions Start--
1. Admin > Advanced > Special Content > Python
2. New Python Script File
3. Name it "TPxi_LeaderRetention" and paste this code
4. Open it at /PyScriptForm/TPxi_LeaderRetention
The first line (#roles=Admin) decides who can open it. Change it to, for
example, #roles=Admin,Edit to let more staff in.
--Upload Instructions End--

Written By: Ben Swaby (TPxi Software, LLC)
Email: bswaby@fbchtn.org
Website: https://tpxisoftware.com
GitHub: https://github.com/bswaby/Touchpoint  (50+ free tools)
----------------------------------------------------------------
These tools are free because they should be.
If they've saved you time or helped your team, and you want to
support continued development, check out:

DisplayCache(TM) - church digital signage that integrates with TouchPoint(R)
https://displaycache.com

TPxi Go(TM) - your church contacts, wherever you work.
Look up anyone in TouchPoint(R), log calls and emails from Outlook
or your phone. No tab switching, no lost context.
https://tpxigo.com
----------------------------------------------------------------

Version: 1.1
"""

import json
import datetime
import traceback

PAGE_TITLE = "Leader Retention"
SETTINGS_KEY = "TPxi_LeaderRetention_Settings"

# A row with fewer leaders than this still shows, but its percentages are
# greyed out: two leaders out of three is not a trend.
SMALL_N = 10

# Roster basis: days in a leader role within the year to count as leading.
ROSTER_MIN_DAYS = 30

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]


#####################################################################
#### HELPERS
#####################################################################

def get_param(name, default=""):
    try:
        v = getattr(model.Data, name)
        if v is None:
            return default
        v = str(v).strip()
        return v if v else default
    except Exception:
        return default


def h(s):
    if s is None:
        return ""
    s = unicode(s) if not isinstance(s, (str, unicode)) else s
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def to_int(s, default):
    try:
        return int(str(s).strip())
    except Exception:
        return default


def id_list(s):
    out = []
    for part in str(s or "").split(","):
        part = part.strip()
        if part.isdigit():
            out.append(int(part))
    return out


def sql_ids(ids):
    return ",".join(str(int(i)) for i in ids) or "0"


#####################################################################
#### SETTINGS (saved on the page, not edited in the script)
#####################################################################

def leader_role_options():
    """Member types TouchPoint counts as leader or volunteer attendance."""
    return list(q.QuerySql("""
        SELECT Id, Description, AttendanceTypeId
        FROM lookup.MemberType
        WHERE AttendanceTypeId IN (10, 20)
        ORDER BY AttendanceTypeId, Description
    """))


def default_settings():
    roles = [int(r.Id) for r in leader_role_options()
             if int(r.AttendanceTypeId or 0) == 10
             and "staff" not in (r.Description or "").lower()]
    return {"basis": "attended", "fy_start": 10, "years": 3,
            "roles": roles, "excluded": []}


def load_settings():
    s = default_settings()
    try:
        raw = model.TextContent(SETTINGS_KEY)
        if raw:
            saved = json.loads(raw)
            for k in s:
                if k in saved:
                    s[k] = saved[k]
    except Exception:
        pass
    return s


def save_settings(s):
    # Ids and short codes only, so plain json is safe here.
    model.WriteContentText(SETTINGS_KEY, json.dumps(s), "")


def effective_settings(saved):
    """Saved defaults, overridden for this view by anything in the address."""
    s = dict(saved)
    basis = get_param("basis")
    if basis in ("attended", "roster"):
        s["basis"] = basis
    m = to_int(get_param("fystart"), 0)
    if 1 <= m <= 12:
        s["fy_start"] = m
    y = to_int(get_param("years"), 0)
    if 1 <= y <= 6:
        s["years"] = y
    s["fy_start"] = to_int(s.get("fy_start"), 10)
    s["years"] = to_int(s.get("years"), 3)
    s["roles"] = [int(x) for x in s.get("roles") or []]
    s["excluded"] = [int(x) for x in s.get("excluded") or []]
    return s


#####################################################################
#### FISCAL YEARS
#####################################################################
# A fiscal year is named for the calendar year it ENDS in. With an October
# start, FY2026 runs Oct 1 2025 to Sep 30 2026.

def fy_of(d, start_month):
    if start_month == 1:
        return d.year
    return d.year + 1 if d.month >= start_month else d.year


def fy_bounds(fy, start_month):
    first = datetime.date(fy - (0 if start_month == 1 else 1), start_month, 1)
    nxt = datetime.date(first.year + 1, first.month, 1)
    return first, nxt - datetime.timedelta(days=1)


def fy_shift_sql(start_month):
    """Months to add so YEAR() of the result is the fiscal year."""
    return (13 - start_month) % 12


#####################################################################
#### DATA
#####################################################################

def fy_windows(s, fys, today):
    """(label, first day, last day) for each fiscal year, cut off at today."""
    out = []
    for fy in fys:
        a, b = fy_bounds(fy, s["fy_start"])
        out.append((fy, a, min(b, today)))
    return out


def load_leader_rows(s, windows):
    """(FY, PeopleId, ProgId, DivId) for everyone who led in each window."""
    roles = sql_ids(s["roles"])
    excl = sql_ids(s["excluded"])
    values = []
    for label, a, b in windows:
        need = min(ROSTER_MIN_DAYS, max((b - a).days, 1))
        values.append("({0}, '{1}', '{2}', {3})".format(label, a.isoformat(), b.isoformat(), need))
    lo = min(w[1] for w in windows).isoformat()
    hi = max(w[2] for w in windows).isoformat()

    if s["basis"] == "attended":
        sql = """
            ;WITH fy AS (SELECT * FROM (VALUES {0}) v(FY, s, e, need))
            SELECT DISTINCT fy.FY, a.PeopleId, d.ProgId, o.DivisionId AS DivId
            FROM dbo.Attend a WITH (NOLOCK)
            JOIN dbo.Organizations o ON o.OrganizationId = a.OrganizationId
            JOIN dbo.Division d ON d.Id = o.DivisionId
            JOIN fy ON a.MeetingDate >= fy.s AND a.MeetingDate < DATEADD(day, 1, fy.e)
            WHERE a.AttendanceFlag = 1 AND a.MemberTypeId IN ({1})
              AND a.MeetingDate >= '{3}' AND a.MeetingDate < DATEADD(day, 1, '{4}')
              AND d.ProgId NOT IN ({2})
        """.format(",".join(values), roles, excl, lo, hi)
    else:
        sql = """
            ;WITH fy AS (SELECT * FROM (VALUES {0}) v(FY, s, e, need))
            SELECT DISTINCT fy.FY, et.PeopleId, d.ProgId, o.DivisionId AS DivId
            FROM dbo.EnrollmentTransaction et WITH (NOLOCK)
            JOIN dbo.Organizations o ON o.OrganizationId = et.OrganizationId
            JOIN dbo.Division d ON d.Id = o.DivisionId
            JOIN fy ON et.TransactionDate <= fy.e
                   AND ISNULL(et.NextTranChangeDate, GETDATE()) >= fy.s
            WHERE et.TransactionTypeId IN (1, 3) AND et.MemberTypeId IN ({1})
              AND d.ProgId NOT IN ({2})
              AND DATEDIFF(day,
                    CASE WHEN et.TransactionDate > fy.s THEN et.TransactionDate ELSE fy.s END,
                    CASE WHEN ISNULL(et.NextTranChangeDate, GETDATE()) < fy.e
                         THEN ISNULL(et.NextTranChangeDate, GETDATE()) ELSE fy.e END) >= fy.need
        """.format(",".join(values), roles, excl)
    return q.QuerySql(sql)


class Sets(object):
    """Who led where, per fiscal year."""
    def __init__(self, rows):
        self.prog = {}
        self.div = {}
        self.any = {}
        self.p_progs = {}
        self.p_divs = {}
        for r in rows:
            fy, pid = int(r.FY), int(r.PeopleId)
            pg, dv = int(r.ProgId or 0), int(r.DivId or 0)
            self.prog.setdefault((fy, pg), set()).add(pid)
            self.div.setdefault((fy, dv), set()).add(pid)
            self.any.setdefault(fy, set()).add(pid)
            self.p_progs.setdefault((fy, pid), set()).add(pg)
            self.p_divs.setdefault((fy, pid), set()).add(dv)

    def scope_set(self, scope, fy):
        kind, sid = scope
        if kind == "p":
            return self.prog.get((fy, sid), set())
        if kind == "d":
            return self.div.get((fy, sid), set())
        return self.any.get(fy, set())


def split(sets, scope, fy):
    """Leaders in scope in fy, split by what they did in fy + 1."""
    base = sets.scope_set(scope, fy)
    nxt_scope = sets.scope_set(scope, fy + 1)
    nxt_any = sets.any.get(fy + 1, set())
    kept = base & nxt_scope
    anywhere = base & nxt_any
    return {"base": base, "kept": kept, "moved": anywhere - kept,
            "stopped": base - anywhere}


def load_names():
    progs, divs = {}, {}
    for r in q.QuerySql("SELECT Id, Name FROM dbo.Program"):
        progs[int(r.Id)] = r.Name or ("Program %d" % int(r.Id))
    for r in q.QuerySql("SELECT Id, Name, ProgId FROM dbo.Division"):
        divs[int(r.Id)] = (r.Name or ("Division %d" % int(r.Id)), int(r.ProgId or 0))
    return progs, divs


#####################################################################
#### RENDERING: RETENTION TABLE
#####################################################################

def pct(n, d):
    return (100.0 * n / d) if d else 0.0


def color(p, n):
    if n < SMALL_N:
        return "#999"
    if p >= 80:
        return "#1e8449"
    if p >= 65:
        return "#b9770e"
    return "#c0392b"


def link(params, label, title=""):
    qs = "&amp;".join("{0}={1}".format(k, h(str(v))) for k, v in params)
    return '<a href="?{0}#followup" title="{1}">{2}</a>'.format(qs, h(title), label)


def pace_note(kp, n, pace):
    """Color and a line comparing the year so far with the same point last
    year. A few weeks in, most people who serve monthly or on a rotation
    have not had a turn yet, so the full-year thresholds do not apply."""
    if pace is None or not pace["base"]:
        return "#999", '<br><i>so far; no leaders here last year to compare</i>'
    pp = pct(len(pace["kept"]), len(pace["base"]))
    if pp == 0 and kp == 0:
        return "#999", ('<br><i>so far; none by this point last year either</i>'
                        '<br><b style="color:#999">usually starts later</b>')
    if n < SMALL_N or len(pace["base"]) < SMALL_N:
        clr, word = "#999", "small group"
    elif kp >= pp - 5:
        clr, word = "#1e8449", "on pace"
    elif kp >= pp - 15:
        clr, word = "#b9770e", "a little behind"
    else:
        clr, word = "#c0392b", "behind"
    return clr, ('<br><i>so far; this point last year: {0:.0f}%</i>'
                 '<br><b style="color:{1}">{2}</b>'.format(pp, clr, word))


def cell(sp, scope_key, fy, keep_params, in_progress=False, pace=None):
    n = len(sp["base"])
    if n == 0:
        return '<td class="c muted">&ndash;</td>'
    kp = pct(len(sp["kept"]), n)
    ap = pct(len(sp["kept"]) + len(sp["moved"]), n)
    clr = color(kp, n)
    tail = ''
    if in_progress:
        clr, tail = pace_note(kp, n, pace)
    base = keep_params + [("scope", scope_key), ("from", fy)]
    stopped = link(base + [("status", "stopped")], "%d stopped" % len(sp["stopped"]),
                   "Led here, not leading anywhere the next year") if sp["stopped"] else "0 stopped"
    moved = link(base + [("status", "moved")], "%d moved" % len(sp["moved"]),
                 "Led here, leading somewhere else the next year") if sp["moved"] else "0 moved"
    if scope_key == "all":
        # Church-wide nobody can move elsewhere, so kept and still leading
        # are the same number.
        return ('<td class="c"><div class="big" style="color:{0}">{1:.0f}%</div>'
                '<div class="sub">{2} of {3} still leading<br>{4}{5}</div></td>').format(
                    clr, kp, len(sp["kept"]), n, stopped, tail)
    return ('<td class="c"><div class="big" style="color:{0}">{1:.0f}%</div>'
            '<div class="sub">kept {2} of {3}<br>{4:.0f}% still leading<br>{5} &middot; {6}{7}</div>'
            '</td>').format(clr, kp, len(sp["kept"]), n, ap, stopped, moved, tail)


def render_table(sets, s, pairs, progs, divs, keep_params, cur_fy, pace_sets=None):
    def pc(scope, fy):
        if fy + 1 != cur_fy or pace_sets is None:
            return None
        return split(pace_sets, scope, cur_fy - 2)

    out = ['<div class="scroll"><table class="lr"><thead><tr><th class="name">Program'
           ' <button type="button" class="mini" onclick="lrAll(true)">expand all</button>'
           ' <button type="button" class="mini" onclick="lrAll(false)">collapse</button></th>']
    for fy in pairs:
        label = "FY{0} &rarr; FY{1}".format(str(fy)[2:], str(fy + 1)[2:])
        if fy + 1 == cur_fy:
            label += '<div class="sub">year in progress</div>'
        out.append('<th class="c">{0}</th>'.format(label))
    out.append('</tr></thead><tbody>')

    # church-wide
    out.append('<tr class="total"><td class="name"><b>All programs</b>'
               '<div class="sub">anyone who led anywhere</div></td>')
    for fy in pairs:
        out.append(cell(split(sets, ("all", 0), fy), "all", fy, keep_params, fy + 1 == cur_fy,
                        pc(("all", 0), fy)))
    out.append('</tr>')

    prog_ids = set()
    div_ids = {}
    for fy in pairs:
        for (f, pg) in sets.prog:
            if f == fy:
                prog_ids.add(pg)
        for (f, dv) in sets.div:
            if f == fy:
                div_ids.setdefault(divs.get(dv, ("", 0))[1], set()).add(dv)

    def biggest(pg):
        return -max([len(sets.prog.get((fy, pg), ())) for fy in pairs] or [0])

    for pg in sorted(prog_ids, key=lambda p: (biggest(p), progs.get(p, ""))):
        kids = sorted(div_ids.get(pg, set()), key=lambda d: divs.get(d, ("",))[0])
        tog = ('<span class="tog" id="tog-{0}">&#9656;</span> '.format(pg) if kids else
               '<span class="tog"></span> ')
        out.append('<tr class="prog" data-prog="{0}" onclick="lrToggle({0})">'
                   '<td class="name">{1}<b>{2}</b><div class="sub">{3} division{4}</div></td>'.format(
                       pg, tog, h(progs.get(pg, "Program %d" % pg)), len(kids),
                       "" if len(kids) == 1 else "s"))
        for fy in pairs:
            out.append(cell(split(sets, ("p", pg), fy), "p%d" % pg, fy, keep_params, fy + 1 == cur_fy,
                            pc(("p", pg), fy)))
        out.append('</tr>')
        for dv in kids:
            out.append('<tr class="div div-{0}" style="display:none"><td class="name">'
                       '<span class="indent"></span>{1}</td>'.format(pg, h(divs.get(dv, ("Division %d" % dv,))[0])))
            for fy in pairs:
                out.append(cell(split(sets, ("d", dv), fy), "d%d" % dv, fy, keep_params, fy + 1 == cur_fy,
                                pc(("d", dv), fy)))
            out.append('</tr>')
    out.append('</tbody></table></div>')
    return "".join(out)


#####################################################################
#### RENDERING: FOLLOW-UP LIST
#####################################################################

def parse_scope(raw):
    raw = (raw or "").strip()
    if raw == "all":
        return ("all", 0)
    if len(raw) > 1 and raw[0] in "pd" and raw[1:].isdigit():
        return (raw[0], int(raw[1:]))
    return None


def scope_label(scope, progs, divs):
    kind, sid = scope
    if kind == "p":
        return progs.get(sid, "Program %d" % sid)
    if kind == "d":
        name, pg = divs.get(sid, ("Division %d" % sid, 0))
        return "{0} / {1}".format(progs.get(pg, ""), name)
    return "All programs"


def scope_sql(scope, s):
    kind, sid = scope
    excl = "d.ProgId NOT IN ({0})".format(sql_ids(s["excluded"]))
    if kind == "p":
        return "d.ProgId = {0}".format(sid)
    if kind == "d":
        return "o.DivisionId = {0} AND {1}".format(sid, excl)
    return excl


def followup_rows(pids, scope, fy, s, today):
    if not pids:
        return []
    roles = sql_ids(s["roles"])
    a, b = fy_bounds(fy, s["fy_start"])
    na, nb = fy_bounds(fy + 1, s["fy_start"])
    nb = min(nb, today)
    where = scope_sql(scope, s)
    out = []
    ids = sorted(pids)
    for i in range(0, len(ids), 500):
        chunk = sql_ids(ids[i:i + 500])
        sql = """
            SELECT p.PeopleId, p.Name, p.Name2, p.Age, p.EmailAddress, p.CellPhone,
                p.IsDeceased, ms.Description AS MemberStatus,
                lead.OrganizationId AS LastOrgId, lead.OrganizationName AS LastOrg,
                CONVERT(varchar(10), lead.MeetingDate, 120) AS LastLed,
                ISNULL(cnt.N, 0) AS TimesLed,
                ros.OrganizationId AS RosterOrgId, ros.OrganizationName AS RosterOrg,
                ISNULL(nx.N, 0) AS NextAttends, CONVERT(varchar(10), nx.LastDt, 120) AS NextLast,
                still.Orgs AS StillListed
            FROM dbo.People p
            LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
            OUTER APPLY (
                SELECT TOP 1 a.OrganizationId, o.OrganizationName, a.MeetingDate
                FROM dbo.Attend a WITH (NOLOCK)
                JOIN dbo.Organizations o ON o.OrganizationId = a.OrganizationId
                JOIN dbo.Division d ON d.Id = o.DivisionId
                WHERE a.PeopleId = p.PeopleId AND a.AttendanceFlag = 1
                  AND a.MemberTypeId IN ({1}) AND {2}
                  AND a.MeetingDate >= '{3}' AND a.MeetingDate < DATEADD(day, 1, '{4}')
                ORDER BY a.MeetingDate DESC
            ) lead
            OUTER APPLY (
                SELECT COUNT(*) AS N
                FROM dbo.Attend a WITH (NOLOCK)
                JOIN dbo.Organizations o ON o.OrganizationId = a.OrganizationId
                JOIN dbo.Division d ON d.Id = o.DivisionId
                WHERE a.PeopleId = p.PeopleId AND a.AttendanceFlag = 1
                  AND a.MemberTypeId IN ({1}) AND {2}
                  AND a.MeetingDate >= '{3}' AND a.MeetingDate < DATEADD(day, 1, '{4}')
            ) cnt
            OUTER APPLY (
                SELECT TOP 1 et.OrganizationId, o.OrganizationName
                FROM dbo.EnrollmentTransaction et WITH (NOLOCK)
                JOIN dbo.Organizations o ON o.OrganizationId = et.OrganizationId
                JOIN dbo.Division d ON d.Id = o.DivisionId
                WHERE et.PeopleId = p.PeopleId AND et.TransactionTypeId IN (1, 3)
                  AND et.MemberTypeId IN ({1}) AND {2}
                  AND et.TransactionDate <= '{4}'
                  AND ISNULL(et.NextTranChangeDate, GETDATE()) >= '{3}'
                ORDER BY et.TransactionDate DESC
            ) ros
            OUTER APPLY (
                SELECT COUNT(*) AS N, MAX(a.MeetingDate) AS LastDt
                FROM dbo.Attend a WITH (NOLOCK)
                WHERE a.PeopleId = p.PeopleId AND a.AttendanceFlag = 1
                  AND a.MeetingDate >= '{5}' AND a.MeetingDate < DATEADD(day, 1, '{6}')
            ) nx
            OUTER APPLY (
                SELECT STRING_AGG(CAST(o.OrganizationName AS NVARCHAR(MAX)), '; ') AS Orgs
                FROM dbo.OrganizationMembers om
                JOIN dbo.Organizations o ON o.OrganizationId = om.OrganizationId
                JOIN dbo.Division d ON d.Id = o.DivisionId
                WHERE om.PeopleId = p.PeopleId AND om.MemberTypeId IN ({1}) AND {2}
                  AND o.OrganizationStatusId = 30
            ) still
            WHERE p.PeopleId IN ({0})
        """.format(chunk, roles, where, a.isoformat(), b.isoformat(),
                   na.isoformat(), nb.isoformat())
        out.extend(q.QuerySql(sql))
    return out


def render_followup(sets, scope, fy, status, s, progs, divs, keep_params, today):
    sp = split(sets, scope, fy)
    pids = sp.get(status, set())
    label = scope_label(scope, progs, divs)
    verb = {"stopped": "did not lead anywhere", "moved": "led somewhere else",
            "kept": "led here again"}.get(status, "")
    tabs = []
    for st in ("stopped", "moved", "kept"):
        n = len(sp[st])
        txt = "{0} ({1})".format(st.capitalize(), n)
        if st == status:
            tabs.append('<b class="pill on">{0}</b>'.format(txt))
        else:
            tabs.append(link(keep_params + [("scope", "%s%s" % (scope[0], scope[1] or "")
                                             if scope[0] != "all" else "all"),
                                            ("from", fy), ("status", st)],
                             '<span class="pill">{0}</span>'.format(txt)))

    out = ['<div id="followup" class="panel">',
           '<h3>Follow-up: {0}</h3>'.format(h(label)),
           '<p>Led here in <b>FY{0}</b> and {1} in <b>FY{2}</b>{3}. {4}</p>'.format(
               fy, verb, fy + 1,
               " (so far)" if fy_of(today, s["fy_start"]) == fy + 1 else "",
               " ".join(tabs))]
    rows = followup_rows(pids, scope, fy, s, today)
    if not rows:
        out.append('<p class="muted">Nobody in this group.</p></div>')
        return "".join(out)

    def sort_key(r):
        # Still attending but not leading first: the easiest to re-engage.
        return (int(r.IsDeceased or 0), 0 if int(r.NextAttends or 0) else 1,
                -int(r.TimesLed or 0), r.Name2 or "")
    rows = sorted(rows, key=sort_key)

    out.append('<div class="actions"><button type="button" class="btn" onclick="lrCsv()">'
               'Download CSV</button> <button type="button" class="btn" onclick="lrPrint()">'
               'Print</button></div>')
    out.append('<div class="scroll"><table class="lr fu" id="fuTable"><thead><tr>'
               '<th>Name</th><th class="c">Age</th><th>Last led here</th><th class="c">Times led</th>'
               '<th>FY{0}</th><th>Still on a leader roster here</th><th>Member status</th>'
               '<th>Email</th><th>Cell</th></tr></thead><tbody>'.format(fy + 1))
    for r in rows:
        pid = int(r.PeopleId)
        if r.LastOrg:
            last = '<a href="/Org/{0}" target="_blank">{1}</a><div class="sub">{2}</div>'.format(
                int(r.LastOrgId), h(r.LastOrg), h(r.LastLed))
        elif r.RosterOrg:
            last = '<a href="/Org/{0}" target="_blank">{1}</a><div class="sub">on roster, no leader attendance</div>'.format(
                int(r.RosterOrgId), h(r.RosterOrg))
        else:
            last = '<span class="muted">&ndash;</span>'

        if int(r.IsDeceased or 0):
            nxt = '<span class="muted">Deceased</span>'
        elif status == "moved":
            dest = sorted(progs.get(pg, "") for pg in sets.p_progs.get((fy + 1, pid), ()))
            if scope[0] == "d":
                dest = sorted("{0} / {1}".format(progs.get(divs.get(d, ("", 0))[1], ""),
                                                 divs.get(d, ("",))[0])
                              for d in sets.p_divs.get((fy + 1, pid), ()))
            nxt = "Now leads in: " + h(", ".join(dest))
        elif int(r.NextAttends or 0):
            n = int(r.NextAttends)
            nxt = ('<span class="ok">Still attends</span> ({0} time{1}, last {2})'.format(
                n, "" if n == 1 else "s", h(r.NextLast)))
        else:
            nxt = '<span class="warn">No attendance</span>'

        out.append('<tr><td><a href="/Person2/{0}" target="_blank"><b>{1}</b></a></td>'
                   '<td class="c">{2}</td><td>{3}</td><td class="c">{4}</td><td>{5}</td>'
                   '<td class="wrap">{6}</td><td>{7}</td><td>{8}</td><td style="white-space:nowrap">{9}</td></tr>'.format(
                       pid, h(r.Name), int(r.Age) if r.Age is not None else "", last,
                       int(r.TimesLed or 0), nxt, h(r.StillListed or ""), h(r.MemberStatus or ""),
                       ('<a href="mailto:{0}">{0}</a>'.format(h(r.EmailAddress)) if r.EmailAddress else ""),
                       h(model.FmtPhone(r.CellPhone, "") if r.CellPhone else "")))
    out.append('</tbody></table></div>')
    out.append('<p class="note">Sorted with people who still attend first, since they are '
               'the easiest to invite back. &ldquo;Still on a leader roster here&rdquo; means '
               'TouchPoint still lists them as a leader in this area even though they did not '
               'lead; worth updating the roster either way.</p></div>')
    return "".join(out)


#####################################################################
#### RENDERING: SETTINGS
#####################################################################

def render_settings(saved, s, progs, msg):
    roles = leader_role_options()
    role_boxes = []
    for r in roles:
        rid = int(r.Id)
        kind = "leader" if int(r.AttendanceTypeId or 0) == 10 else "volunteer"
        role_boxes.append('<label class="chk"><input type="checkbox" class="lr-role" value="{0}"{1}> '
                          '{2} <span class="muted">({3})</span></label>'.format(
                              rid, " checked" if rid in saved["roles"] else "", h(r.Description), kind))
    prog_boxes = []
    for pid, name in sorted(progs.items(), key=lambda kv: kv[1]):
        prog_boxes.append('<label class="chk"><input type="checkbox" class="lr-excl" value="{0}"{1}> '
                          '{2}</label>'.format(pid, " checked" if pid in saved["excluded"] else "", h(name)))
    month_opts = "".join('<option value="{0}"{1}>{2}</option>'.format(
        i + 1, " selected" if saved["fy_start"] == i + 1 else "", m) for i, m in enumerate(MONTHS))
    year_opts = "".join('<option value="{0}"{1}>{0}</option>'.format(
        n, " selected" if saved["years"] == n else "") for n in range(1, 7))

    return """
    <details class="panel"{open}><summary><b>Settings</b> <span class="muted">
      leader roles, fiscal year, programs to leave out</span></summary>
      {msg}
      <form method="post" action="" onsubmit="return lrSave(this)">
        <input type="hidden" name="lr_action" value="save">
        <input type="hidden" name="lr_roles" value="">
        <input type="hidden" name="lr_excl" value="">
        <div class="grid">
          <div><label class="lbl">Count someone as leading when they</label>
            <select name="lr_basis">
              <option value="attended"{att}>were marked present in a leader role</option>
              <option value="roster"{ros}>were on a leader roster 30+ days</option>
            </select>
            <div class="sub">Attendance is the better measure: leaders are rarely taken off
            rosters. Use the roster for areas that do not take attendance.</div></div>
          <div><label class="lbl">Fiscal year starts in</label>
            <select name="lr_fystart">{months}</select></div>
          <div><label class="lbl">Years to compare</label>
            <select name="lr_years">{years}</select></div>
        </div>
        <label class="lbl">Leader roles</label>
        <div class="boxes">{roles}</div>
        <label class="lbl">Leave these programs out</label>
        <div class="sub">Programs where people serve for a week (VBS, camps) make the
        church-wide rate look worse than it is, because most of those volunteers serve once.
        People who only lead in a program left out here count as not leading.</div>
        <div class="boxes">{progs}</div>
        <button class="btn" type="submit">Save as the default for everyone</button>
      </form>
    </details>
    """.format(open=" open" if msg else "", msg=msg,
               att=" selected" if saved["basis"] == "attended" else "",
               ros=" selected" if saved["basis"] == "roster" else "",
               months=month_opts, years=year_opts,
               roles="".join(role_boxes), progs="".join(prog_boxes))


def render_view_bar(s, keep_params, inprog):
    def opt(params, label, on):
        if on:
            return '<b class="pill on">{0}</b>'.format(label)
        return '<a class="pill" href="?{0}">{1}</a>'.format(
            "&amp;".join("{0}={1}".format(k, h(str(v))) for k, v in params), label)
    base = [(k, v) for k, v in keep_params if k not in ("basis", "inprog")]
    return ('<div class="bar">Counting leaders who {0} {1} &nbsp;&middot;&nbsp; {2}</div>'.format(
        opt(base + [("basis", "attended"), ("inprog", 1 if inprog else 0)],
            "attended as a leader", s["basis"] == "attended"),
        opt(base + [("basis", "roster"), ("inprog", 1 if inprog else 0)],
            "were on a leader roster", s["basis"] == "roster"),
        opt(base + [("basis", s["basis"]), ("inprog", 0 if inprog else 1)],
            ("Hide" if inprog else "Show") + " the year in progress", False)))


#####################################################################
#### PAGE
#####################################################################

CSS = """
<style>
.lrw { font-family:'Segoe UI',Arial,sans-serif; color:#222; max-width:1300px; }
.lrw .note, .lrw .sub { font-size:11px; color:#777; line-height:1.4; }
.lrw .muted { color:#999; }
.lrw .intro { background:#f3f8ff; border-left:4px solid #2a72d4; padding:10px 14px;
              font-size:13px; line-height:1.5; border-radius:4px; margin:10px 0; }
.lrw .panel { border:1px solid #e0e0e0; border-radius:6px; padding:10px 14px; margin:12px 0; }
.lrw .bar { margin:10px 0; font-size:13px; }
.lrw .pill { display:inline-block; padding:3px 10px; border:1px solid #ccd; border-radius:12px;
             margin:2px; text-decoration:none; color:#2c5aa0; font-size:12px; }
.lrw .pill.on { background:#2c5aa0; color:#fff; border-color:#2c5aa0; }
.lrw .scroll { overflow-x:auto; }
.lrw table.lr { border-collapse:collapse; font-size:12px; margin-top:6px; }
.lrw table.lr th, .lrw table.lr td { border:1px solid #ddd; padding:6px 8px; vertical-align:top; }
.lrw table.lr th { background:#f0f0f0; text-align:left; }
.lrw table.lr .c { text-align:center; }
.lrw table.lr td.name { min-width:220px; }
.lrw table.lr .big { font-size:18px; font-weight:bold; }
.lrw tr.prog { cursor:pointer; }
.lrw tr.prog:hover td { background:#f7fbff; }
.lrw tr.total td { background:#eef4fb; }
.lrw tr.div td { background:#fafafa; }
.lrw .tog { display:inline-block; width:14px; color:#2c5aa0; }
.lrw .indent { display:inline-block; width:22px; }
.lrw .ok { color:#1e8449; font-weight:bold; }
.lrw .warn { color:#c0392b; font-weight:bold; }
.lrw .btn { background:#2c5aa0; color:#fff; border:0; border-radius:4px; padding:6px 12px;
            cursor:pointer; font-size:12px; }
.lrw .mini { font-size:10px; padding:1px 6px; margin-left:4px; cursor:pointer; }
.lrw .grid { display:flex; flex-wrap:wrap; gap:18px; margin:8px 0; }
.lrw .grid > div { min-width:220px; max-width:360px; }
.lrw .lbl { display:block; font-weight:bold; font-size:12px; margin:8px 0 4px; }
.lrw .boxes { display:flex; flex-wrap:wrap; gap:4px 14px; margin:6px 0 10px; font-size:12px; }
.lrw .chk { white-space:nowrap; }
.lrw table.fu td.wrap { max-width:260px; }
.lrw .actions { margin:6px 0; }
</style>
"""

JS = """
<script>
function lrToggle(pg) {
  var rows = document.querySelectorAll('.div-' + pg);
  if (!rows.length) return;
  var show = rows[0].style.display === 'none';
  for (var i = 0; i < rows.length; i++) rows[i].style.display = show ? '' : 'none';
  var t = document.getElementById('tog-' + pg);
  if (t) t.innerHTML = show ? '&#9662;' : '&#9656;';
}
function lrAll(show) {
  var rows = document.querySelectorAll('tr.div');
  for (var i = 0; i < rows.length; i++) rows[i].style.display = show ? '' : 'none';
  var t = document.querySelectorAll('.tog');
  for (var j = 0; j < t.length; j++) if (t[j].id) t[j].innerHTML = show ? '&#9662;' : '&#9656;';
  if (window.event) window.event.stopPropagation();
}
function lrSave(f) {
  var pick = function (cls) {
    var out = [], els = document.querySelectorAll('.' + cls);
    for (var i = 0; i < els.length; i++) if (els[i].checked) out.push(els[i].value);
    return out.join(',');
  };
  f.lr_roles.value = pick('lr-role');
  f.lr_excl.value = pick('lr-excl');
  if (!f.lr_roles.value) { alert('Pick at least one leader role.'); return false; }
  return true;
}
function lrRows() {
  var t = document.getElementById('fuTable');
  if (!t) return [];
  var out = [];
  for (var i = 0; i < t.rows.length; i++) {
    var r = [];
    for (var j = 0; j < t.rows[i].cells.length; j++)
      r.push(t.rows[i].cells[j].innerText.replace(/\\s+/g, ' ').trim());
    out.push(r);
  }
  return out;
}
function lrCsv() {
  var csv = lrRows().map(function (r) {
    return r.map(function (c) { return '"' + c.replace(/"/g, '""') + '"'; }).join(',');
  }).join('\\r\\n');
  var a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], {type: 'text/csv'}));
  a.download = 'leader-followup.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}
function lrPrint() {
  var panel = document.getElementById('followup');
  if (!panel) return;
  var css = '*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}'
    + 'body{font-family:Segoe UI,Arial,sans-serif;font-size:11px;color:#222;padding:16px}'
    + 'table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:4px 6px;text-align:left;vertical-align:top}'
    + 'th{background:#f0f0f0}.sub{font-size:10px;color:#777}.ok{color:#1e8449}.warn{color:#c0392b}'
    + '.actions,.pill{display:none}a{color:#222;text-decoration:none}';
  var pw = window.open('', '_blank');
  if (!pw) { alert('Popup blocked. Please allow popups for this site.'); return; }
  pw.document.write('<!DOCTYPE html><html><head><title>Leader follow-up</title><style>' + css
    + '</style></head><body>' + panel.innerHTML + '</body></html>');
  pw.document.close(); pw.focus();
  setTimeout(function () { pw.print(); }, 300);
}
</script>
"""


def handle_save(saved):
    if str(model.HttpMethod).lower() != "post" or get_param("lr_action") != "save":
        return saved, ""
    roles = id_list(get_param("lr_roles"))
    if not roles:
        return saved, '<p class="warn">Pick at least one leader role. Nothing was saved.</p>'
    new = {"basis": get_param("lr_basis") if get_param("lr_basis") in ("attended", "roster") else "attended",
           "fy_start": min(max(to_int(get_param("lr_fystart"), 10), 1), 12),
           "years": min(max(to_int(get_param("lr_years"), 3), 1), 6),
           "roles": roles,
           "excluded": id_list(get_param("lr_excl"))}
    save_settings(new)
    return new, '<p class="ok">Saved. These are now the defaults for everyone who opens this page.</p>'


def main():
    model.Header = PAGE_TITLE
    today = datetime.date.today()
    saved, msg = handle_save(load_settings())
    s = effective_settings(saved)
    inprog = get_param("inprog") == "1"
    cur_fy = fy_of(today, s["fy_start"])

    # Pairs are named by their starting year. The latest complete pair ends
    # with the fiscal year that just finished.
    last_from = cur_fy - 2
    pairs = list(range(last_from - s["years"] + 1, last_from + 1))
    if inprog:
        pairs.append(cur_fy - 1)
    fys = sorted(set(pairs + [p + 1 for p in pairs]))

    keep_params = [("basis", s["basis"]), ("fystart", s["fy_start"]),
                   ("years", s["years"]), ("inprog", 1 if inprog else 0)]

    progs, divs = load_names()
    sets = Sets(load_leader_rows(s, fy_windows(s, fys, today)))

    # Benchmark for the year in progress: the year before last, and how many
    # of its leaders had led again by this same point last year.
    pace_sets = None
    if inprog:
        start_cur, _ = fy_bounds(cur_fy, s["fy_start"])
        elapsed = today - start_cur
        start_prev, end_prev = fy_bounds(cur_fy - 1, s["fy_start"])
        pace_sets = Sets(load_leader_rows(s, fy_windows(s, [cur_fy - 2], today) +
                                          [(cur_fy - 1, start_prev, min(start_prev + elapsed, end_prev))]))

    html = ['<div class="lrw">', CSS, JS, '<h2>{0}</h2>'.format(PAGE_TITLE)]
    a, _ = fy_bounds(cur_fy, s["fy_start"])
    html.append(
        '<div class="intro">How many of each year&rsquo;s leaders were leading again the next '
        'fiscal year. <b>Kept</b> means the same program (or division), <b>still leading</b> '
        'counts anyone leading anywhere. Click a program to expand its primary divisions, and '
        'click <b>stopped</b> or <b>moved</b> for the follow-up list. Fiscal years start in '
        '{0}; FY{1} began {2}.</div>'.format(MONTHS[s["fy_start"] - 1], cur_fy, a.strftime("%b %d, %Y")))
    html.append(render_view_bar(s, keep_params, inprog))

    # A sudden jump in leader counts usually means roles started being
    # recorded then, which makes the years before look like perfect retention.
    counts = [(fy, len(sets.any.get(fy, ()))) for fy in fys if fy < cur_fy]
    if len(counts) >= 2:
        top = max(c for _, c in counts)
        thin = [fy for fy, c in counts if top and c < top * 0.5]
        if thin:
            html.append('<p class="note" style="color:#b9770e">FY{0} has under half as many '
                        'leaders recorded as later years, so leader roles were probably not '
                        'recorded consistently yet. Treat comparisons from that year with '
                        'caution.</p>'.format(", FY".join(str(f) for f in thin)))

    html.append(render_table(sets, s, pairs, progs, divs, keep_params, cur_fy, pace_sets))
    if inprog:
        html.append('<p class="note">The year in progress is colored by <b>pace</b>, not the '
                    'full-year scale: how this year so far compares with the same number of days '
                    'into last year. Early in the year most people who serve monthly or on a '
                    'rotation have not had a turn yet, and programs that start later (VBS, fall '
                    'classes) show 0% until they meet.</p>')
    html.append('<p class="note">Green is 80% kept or better, amber 65 to 79%, red below 65%. '
                'Grey means fewer than {0} leaders, too few to read much into. Leader counts '
                'by year: {1}.</p>'.format(SMALL_N, ", ".join(
                    "FY{0} {1}".format(str(fy)[2:], len(sets.any.get(fy, ()))) for fy in fys)))

    scope = parse_scope(get_param("scope"))
    fy_sel = to_int(get_param("from"), 0)
    status = get_param("status", "stopped")
    if status not in ("stopped", "moved", "kept"):
        status = "stopped"
    if scope and fy_sel in pairs:
        html.append(render_followup(sets, scope, fy_sel, status, s, progs, divs, keep_params, today))

    html.append(render_settings(saved, s, progs, msg))
    html.append('<div class="note" style="margin-top:20px;border-top:1px solid #eee;padding-top:8px">'
                'Leader Retention &middot; part of a free TouchPoint&reg; tool set by '
                '<a href="https://tpxisoftware.com" target="_blank" rel="noopener noreferrer">TPxi Software</a>. '
                '<a href="https://github.com/bswaby/Touchpoint" target="_blank" rel="noopener noreferrer">'
                '50+ more on GitHub</a>.</div>')
    html.append('</div>')
    model.Form = "".join(html)


try:
    main()
except Exception as _e:
    model.Header = PAGE_TITLE
    model.Form = ('<div style="font-family:Segoe UI,Arial,sans-serif"><h2>{0}</h2>'
                  '<div style="background:#fff0f0;border-left:4px solid #c0392b;padding:12px 16px">'
                  '<b>The page hit an error.</b><pre style="white-space:pre-wrap;font-size:12px">'
                  '{1}\n\n{2}</pre></div></div>').format(PAGE_TITLE, h(str(_e)), h(traceback.format_exc()))
