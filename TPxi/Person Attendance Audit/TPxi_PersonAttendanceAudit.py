#roles=Admin
# -*- coding: utf-8 -*-
#####################################################################
# TPxi_PersonAttendanceAudit  --  READ-ONLY safe-church audit
#####################################################################
# Person-centric co-attendance audit. Point it at ONE person and it
# reconstructs, from actual attendance (AttendanceFlag=1):
#   1. Every session that person was PRESENT for (org, meeting, date, role).
#   2. Everyone else present at those same sessions (co-attendees), with
#      each person's AGE AT THE TIME of that meeting and a minor flag.
#   3. A minors-detail roll-up: each distinct minor co-present, age range,
#      how many sessions, and the dates.
#
# Read-only (SELECT into temp tables only). Access-restricted (#roles=Admin).
# Intended for authorized reviews only

#####################################################################

import re
import datetime
import traceback

model.Header = "Person Attendance Audit (read-only)"

MINOR_AGE = 18    # under 18 at time of attendance = minor (matches ACS script)
CAP = 40000       # safety cap on co-attendance rows


def _param(name, default=''):
    try:
        v = getattr(model.Data, name)
        v = str(v).strip()
        return v if v else default
    except:
        return default


def _isodate(s):
    return s if re.match(r'^\d{4}-\d{2}-\d{2}$', s or '') else ''


PID = _param('pid', '')
QSEARCH = _param('q', '')
CO = _param('co', '')
try:
    PID = int(PID) if PID else 0
except:
    PID = 0
try:
    CO = int(CO) if CO else 0
except:
    CO = 0

DAYS = _param('days', '')
try:
    DAYS = int(DAYS) if DAYS else 0
except:
    DAYS = 0
FROM = _isodate(_param('from', ''))
TO = _isodate(_param('to', ''))

if FROM or TO:
    WINDOW_LABEL = "%s to %s" % (FROM or 'start', TO or 'now')
elif DAYS:
    WINDOW_LABEL = "last %d days" % DAYS
else:
    WINDOW_LABEL = "all history"


def date_pred(col):
    if FROM and TO:
        return "%s >= '%s' AND %s < DATEADD(DAY, 1, '%s')" % (col, FROM, col, TO)
    if FROM:
        return "%s >= '%s'" % (col, FROM)
    if TO:
        return "%s < DATEADD(DAY, 1, '%s')" % (col, TO)
    if DAYS:
        return "%s >= DATEADD(DAY, -%d, GETDATE())" % (col, DAYS)
    return "1 = 1"


# ---- unicode / html helpers (ASCII-safe; preserves chars as entities) -----
def _u(val):
    if val is None:
        return u''
    if isinstance(val, unicode):
        return val
    try:
        return unicode(val)
    except:
        pass
    try:
        s = str(val)
        for enc in ('utf-8', 'cp1252', 'latin-1'):
            try:
                return s.decode(enc)
            except:
                pass
        return u''.join((c if ord(c) < 128 else u'?') for c in s)
    except:
        return u''


def esc(val):
    s = _u(val)
    out = []
    for ch in s:
        o = ord(ch)
        if ch == u'&':
            out.append('&amp;')
        elif ch == u'<':
            out.append('&lt;')
        elif ch == u'>':
            out.append('&gt;')
        elif ch == u'"':
            out.append('&quot;')
        elif ch == u"'":
            out.append('&#39;')
        elif o < 0x20 and ch not in u'\t\n\r':
            out.append(' ')
        elif o < 128:
            out.append(str(ch))
        else:
            out.append('&#%d;' % o)
    return ''.join(out)


_WD = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']


def weekday(iso):
    try:
        y, m, d = int(iso[0:4]), int(iso[5:7]), int(iso[8:10])
        return _WD[datetime.date(y, m, d).weekday()]
    except:
        return ''


# ---------------------------------------------------------------------------
# DATA
# ---------------------------------------------------------------------------
def fetch_target(pid):
    sql = """
    SELECT p.PeopleId, p.Name2 AS Name, p.FamilyId,
           CONVERT(varchar(10), p.BDate, 126) AS DOB,
           DATEDIFF(YEAR, p.BDate, GETDATE())
             - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, p.BDate, GETDATE()), p.BDate) > GETDATE() THEN 1 ELSE 0 END AS AgeNow,
           ISNULL(g.Description, '') AS Gender,
           ISNULL(ms.Description, '') AS MemberStatus
    FROM dbo.People p
    LEFT JOIN lookup.Gender g ON g.Id = p.GenderId
    LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
    WHERE p.PeopleId = %d
    """ % pid
    for r in q.QuerySql(sql):
        return {'PeopleId': r.PeopleId, 'Name': _u(r.Name), 'FamilyId': r.FamilyId or 0,
                'DOB': r.DOB or '', 'AgeNow': r.AgeNow, 'Gender': _u(r.Gender),
                'MemberStatus': _u(r.MemberStatus)}
    return None


def fetch_sessions(pid):
    # every meeting the target was present for, with their role in that org
    sql = """
    SELECT m.MeetingId, m.OrganizationId, ISNULL(o.OrganizationName, '') AS OrganizationName,
           CONVERT(varchar(10), m.MeetingDate, 126) AS MeetingDate,
           ISNULL(mt.Description, '') AS TargetRole
    FROM dbo.Attend a
    JOIN dbo.Meetings m ON m.MeetingId = a.MeetingId AND m.DidNotMeet = 0
    JOIN dbo.Organizations o ON o.OrganizationId = m.OrganizationId
    LEFT JOIN dbo.OrganizationMembers om ON om.OrganizationId = m.OrganizationId AND om.PeopleId = %d
    LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
    WHERE a.PeopleId = %d AND a.AttendanceFlag = 1 AND %s
    """ % (pid, pid, date_pred('m.MeetingDate'))
    out = []
    for r in q.QuerySql(sql):
        out.append({'MeetingId': r.MeetingId, 'OrgId': r.OrganizationId,
                    'Org': _u(r.OrganizationName), 'Date': r.MeetingDate or '',
                    'Role': _u(r.TargetRole)})
    return out


def fetch_co(pid, fam):
    sql = """
    SET NOCOUNT ON;
    IF OBJECT_ID('tempdb..#tm') IS NOT NULL DROP TABLE #tm;
    SELECT DISTINCT a.MeetingId, m.OrganizationId, m.MeetingDate
    INTO #tm
    FROM dbo.Attend a
    JOIN dbo.Meetings m ON m.MeetingId = a.MeetingId AND m.DidNotMeet = 0
    WHERE a.PeopleId = %d AND a.AttendanceFlag = 1 AND %s;

    SELECT TOP %d
        t.MeetingId, t.OrganizationId, ISNULL(o.OrganizationName, '') AS OrganizationName,
        CONVERT(varchar(10), t.MeetingDate, 126) AS MeetingDate,
        a2.PeopleId, p.Name2 AS Name, CONVERT(varchar(10), p.BDate, 126) AS DOB,
        ISNULL(g.Description, '') AS Gender, ISNULL(p.FamilyId, 0) AS FamilyId,
        CASE WHEN p.BDate IS NULL THEN NULL ELSE
             DATEDIFF(YEAR, p.BDate, t.MeetingDate)
               - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, p.BDate, t.MeetingDate), p.BDate) > t.MeetingDate THEN 1 ELSE 0 END
        END AS AgeAtAttend,
        ISNULL(mt.Description, '') AS RoleInGroup,
        CASE WHEN ISNULL(p.FamilyId, 0) = %d AND %d <> 0 THEN 1 ELSE 0 END AS SameFamily
    FROM #tm t
    JOIN dbo.Attend a2 ON a2.MeetingId = t.MeetingId AND a2.AttendanceFlag = 1 AND a2.PeopleId <> %d
    JOIN dbo.Organizations o ON o.OrganizationId = t.OrganizationId
    JOIN dbo.People p ON p.PeopleId = a2.PeopleId
    LEFT JOIN lookup.Gender g ON g.Id = p.GenderId
    LEFT JOIN dbo.OrganizationMembers om ON om.OrganizationId = t.OrganizationId AND om.PeopleId = a2.PeopleId
    LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
    ORDER BY t.MeetingDate DESC, p.Name2;
    """ % (pid, date_pred('m.MeetingDate'), CAP, fam, fam, pid)
    out = []
    for r in q.QuerySql(sql):
        age = None if r.AgeAtAttend is None else int(r.AgeAtAttend)
        out.append({'MeetingId': r.MeetingId, 'OrgId': r.OrganizationId, 'Org': _u(r.OrganizationName),
                    'Date': r.MeetingDate or '', 'PeopleId': r.PeopleId, 'Name': _u(r.Name),
                    'DOB': r.DOB or '', 'Gender': _u(r.Gender), 'FamilyId': r.FamilyId or 0,
                    'Age': age, 'Minor': (age is not None and age < MINOR_AGE),
                    'Role': _u(r.RoleInGroup), 'SameFam': (int(r.SameFamily or 0) == 1)})
    return out


# ---------------------------------------------------------------------------
# STYLE
# ---------------------------------------------------------------------------
CSS = """
<style>
  .paa { font-family:'Segoe UI',-apple-system,Arial,sans-serif; color:#1e293b; max-width:1200px; line-height:1.5; }
  .paa h2 { color:#7f1d1d; margin:0 0 2px; font-size:24px; }
  .paa .lede { color:#64748b; font-size:13px; margin:0 0 14px; max-width:940px; }
  .paa code { background:#eef2f7; padding:1px 5px; border-radius:3px; font-size:12px; }
  .paa .subject { background:#fef2f2; border:1px solid #fecaca; border-left:5px solid #b91c1c; border-radius:10px; padding:12px 16px; margin:0 0 14px; }
  .paa .subject .nm { font-size:20px; font-weight:800; color:#7f1d1d; }
  .paa .subject .meta { font-size:13px; color:#7a4a4a; margin-top:2px; }
  .paa .cards { display:flex; gap:12px; flex-wrap:wrap; margin:0 0 16px; }
  .paa .card { flex:1; min-width:150px; border:1px solid #e2e8f0; border-top:3px solid #7f1d1d; border-radius:10px; padding:12px 14px; background:#fff; }
  .paa .card .lbl { font-size:11px; text-transform:uppercase; letter-spacing:.6px; color:#64748b; font-weight:600; }
  .paa .card .val { font-size:26px; font-weight:800; margin-top:3px; color:#7f1d1d; }
  .paa .card .sub { font-size:11px; color:#94a3b8; }
  .paa .card.minor { border-top-color:#be123c; } .paa .card.minor .val { color:#be123c; }
  .paa .filters { display:flex; flex-wrap:wrap; align-items:flex-end; gap:14px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px 16px; margin:0 0 14px; }
  .paa .filters label { font-size:11px; text-transform:uppercase; letter-spacing:.5px; color:#64748b; font-weight:600; display:flex; flex-direction:column; gap:3px; }
  .paa .filters input { font-size:14px; padding:6px 9px; border:1px solid #cbd5e1; border-radius:6px; }
  .paa .filters .go { background:#7f1d1d; color:#fff; border:0; border-radius:6px; padding:8px 18px; font-size:14px; cursor:pointer; }
  .paa .filters .reset { color:#64748b; font-size:12px; align-self:center; }
  .paa .sec { margin:22px 0 6px; font-size:17px; color:#7f1d1d; font-weight:800; border-bottom:2px solid #7f1d1d; padding-bottom:3px; }
  .paa .sec .c { color:#94a3b8; font-size:13px; font-weight:500; }
  .paa .toolbar { display:flex; align-items:center; gap:14px; flex-wrap:wrap; margin:0 0 8px; }
  .paa .btn { display:inline-block; background:#7f1d1d; color:#fff; padding:7px 14px; border-radius:6px; font-size:13px; border:0; cursor:pointer; }
  .paa .btn.alt { background:#0f766e; }
  .paa .toolbar .chk { display:flex; align-items:center; gap:6px; font-size:13px; color:#334155; }
  .paa .toolbar .count { color:#64748b; font-size:13px; margin-left:auto; }
  .paa table { border-collapse:collapse; width:100%; font-size:13px; margin:2px 0 6px; }
  .paa th { background:#eef2f7; color:#334155; text-align:left; padding:5px 9px; font-size:11px; text-transform:uppercase; letter-spacing:.4px; cursor:pointer; white-space:nowrap; }
  .paa th:hover { background:#e2e8f0; }
  .paa th .arrow { color:#94a3b8; font-weight:400; }
  .paa thead .flt th { background:#fff; padding:3px 6px; cursor:auto; }
  .paa thead .flt input { width:100%; box-sizing:border-box; font-size:12px; padding:3px 5px; border:1px solid #cbd5e1; border-radius:4px; font-weight:400; text-transform:none; letter-spacing:0; }
  .paa td { padding:5px 9px; border-bottom:1px solid #f1f5f9; }
  .paa a { color:#7f1d1d; text-decoration:none; } .paa a:hover { text-decoration:underline; }
  .paa .num { text-align:center; font-variant-numeric:tabular-nums; }
  .paa .minorflag { color:#be123c; font-weight:800; }
  .paa .fam { color:#b45309; font-size:11px; font-weight:700; }
  .paa .muted { color:#94a3b8; font-size:12px; }
  .paa .cap { background:#fff7ed; border:1px solid #fed7aa; border-left:4px solid #ea580c; color:#7a4a00; border-radius:8px; padding:8px 14px; margin:0 0 12px; font-size:12.5px; }
  .paa .search { display:flex; gap:10px; align-items:flex-end; }
  .paa .rlist a { display:block; padding:6px 4px; border-bottom:1px solid #f1f5f9; }
</style>
"""

PRINT_JS = """
<script>
function paaFilter(id){
  var t=document.getElementById(id); if(!t) return;
  var ins=t.querySelectorAll('thead tr.flt input'); var vals=[];
  for(var i=0;i<ins.length;i++) vals.push(ins[i].value.toLowerCase());
  var hide=document.getElementById('paaHideAdult'); hide=hide&&hide.checked;
  var rows=t.tBodies[0].rows, shown=0;
  for(var r=0;r<rows.length;r++){
    var td=rows[r].cells, ok=true;
    for(var c=0;c<vals.length;c++){ if(vals[c] && (td[c].textContent||'').toLowerCase().indexOf(vals[c])<0){ok=false;break;} }
    if(ok && hide && rows[r].getAttribute('data-minor')!=='1') ok=false;
    rows[r].style.display=ok?'':'none'; if(ok) shown++;
  }
  var s=document.getElementById('paaShown'); if(s) s.textContent=shown;
}
var paaDir={};
function paaSort(id,c){
  var t=document.getElementById(id), tb=t.tBodies[0], rows=[].slice.call(tb.rows);
  var dir=paaDir[id+'_'+c]=-(paaDir[id+'_'+c]||1);
  rows.sort(function(a,b){var x=(a.cells[c].textContent||'').trim().toLowerCase(),y=(b.cells[c].textContent||'').trim().toLowerCase();
    var nx=parseFloat(x),ny=parseFloat(y);
    if(!isNaN(nx)&&!isNaN(ny)){return (nx-ny)*dir;}
    if(x<y)return -1*dir; if(x>y)return 1*dir; return 0;});
  for(var i=0;i<rows.length;i++) tb.appendChild(rows[i]);
}
function paaCsvCell(v){ v=(v==null?'':(''+v)); return '"'+v.replace(/"/g,'""')+'"'; }
function paaExport(id, fname){
  var t=document.getElementById(id); if(!t) return;
  var ths=t.tHead.querySelectorAll('tr.hdr th'); var head=[];
  for(var h=0;h<ths.length;h++){ head.push((ths[h].textContent||'').replace(/[\\u2191\\u2193\\u21C5]/g,'').trim()); }
  head.push('PeopleId'); head.push('OrgId');
  var lines=[head.map(paaCsvCell).join(',')];
  var rows=t.tBodies[0].rows;
  for(var r=0;r<rows.length;r++){ if(rows[r].style.display==='none') continue;
    var td=rows[r].cells, arr=[];
    for(var c=0;c<td.length;c++) arr.push((td[c].textContent||'').trim());
    arr.push(rows[r].getAttribute('data-pid')||''); arr.push(rows[r].getAttribute('data-oid')||'');
    lines.push(arr.map(paaCsvCell).join(','));
  }
  var blob=new Blob([lines.join('\\r\\n')],{type:'text/csv;charset=utf-8;'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=fname;
  document.body.appendChild(a); a.click(); setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(a.href);},100);
}
function paaPrint(){
  var node=document.querySelector('.paa'); if(!node){window.print();return;}
  var clone=node.cloneNode(true);
  var kill=clone.querySelectorAll('.btn,.filters,.toolbar .chk,thead tr.flt');
  for(var i=0;i<kill.length;i++){kill[i].parentNode.removeChild(kill[i]);}
  var rows=clone.querySelectorAll('tbody tr'); for(var r=0;r<rows.length;r++){ if(rows[r].style.display==='none') rows[r].parentNode.removeChild(rows[r]); }
  var css="*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}body{font-family:'Segoe UI',Arial,sans-serif;color:#1e293b;padding:20px}h2{color:#7f1d1d}.subject{background:#fef2f2;border-left:5px solid #b91c1c;padding:8px 12px;margin-bottom:10px}.subject .nm{font-size:16px;font-weight:800;color:#7f1d1d}.sec{font-size:14px;color:#7f1d1d;border-bottom:2px solid #7f1d1d;margin:12px 0 4px;font-weight:800}table{border-collapse:collapse;width:100%;font-size:11px}th{background:#eef2f7;text-align:left;padding:4px 8px}td{padding:4px 8px;border-bottom:1px solid #f1f5f9}.minorflag{color:#be123c;font-weight:800}.fam{color:#b45309;font-weight:700}.cards{display:flex;gap:10px;margin-bottom:12px}.card{border:1px solid #e2e8f0;border-top:3px solid #7f1d1d;border-radius:8px;padding:8px 12px}.card .val{font-size:20px;font-weight:800;color:#7f1d1d}.card .lbl{font-size:10px;text-transform:uppercase;color:#64748b}.arrow{display:none}a{color:#1e293b;text-decoration:none}";
  var pw=window.open('','_blank'); if(!pw){alert('Allow popups to print');return;}
  pw.document.write('<!DOCTYPE html><html><head><title>Person Attendance Audit</title><style>'+css+'</style></head><body>'+clone.innerHTML+'</body></html>');
  pw.document.close(); pw.focus(); setTimeout(function(){pw.print();},350);
}
</script>
"""


def window_form(pid):
    return ('<form method="get" class="filters">'
            '<input type="hidden" name="pid" value="%d">'
            '<label>Last N days<input type="number" name="days" min="1" placeholder="all history" value="%s"></label>'
            '<label>From<input type="date" name="from" value="%s"></label>'
            '<label>To<input type="date" name="to" value="%s"></label>'
            '<button type="submit" class="go">Apply window</button>'
            '<a class="reset" href="?pid=%d">reset</a>'
            '</form>' % (pid, (str(DAYS) if DAYS else ''), FROM, TO, pid))


def render_search():
    h = ['<form method="get" class="filters"><div class="search">'
         '<label style="min-width:280px;">Find a person<input type="text" name="q" value="%s" placeholder="last name, first name"></label>'
         '<button type="submit" class="go">Search</button></div></form>' % esc(QSEARCH)]
    if QSEARCH:
        term = QSEARCH.replace("'", "''")
        sql = ("SELECT TOP 50 p.PeopleId, p.Name2 AS Name, CONVERT(varchar(10),p.BDate,126) AS DOB, "
               "ISNULL(ms.Description,'') AS Status FROM dbo.People p "
               "LEFT JOIN lookup.MemberStatus ms ON ms.Id=p.MemberStatusId "
               "WHERE p.Name2 LIKE '%%%s%%' OR p.Name LIKE '%%%s%%' ORDER BY p.Name2" % (term, term))
        rows = list(q.QuerySql(sql))
        if not rows:
            h.append('<div class="muted">No matches for "%s".</div>' % esc(QSEARCH))
        else:
            h.append('<div class="sec">Matches <span class="c">(%d)</span></div><div class="rlist">' % len(rows))
            for r in rows:
                h.append('<a href="?pid=%d">%s <span class="muted">&middot; #%d &middot; DOB %s &middot; %s</span></a>'
                         % (r.PeopleId, esc(r.Name), r.PeopleId, esc(r.DOB or '?'), esc(r.Status)))
            h.append('</div>')
    else:
        h.append('<div class="muted">Enter a name to find the person to audit, or pass <code>?pid=</code> directly.</div>')
    return ''.join(h)


def subject_block(t):
    dob = t['DOB'] or '?'
    age = ('%d' % t['AgeNow']) if t['AgeNow'] is not None else '?'
    return ('<div class="subject"><div class="nm">%s</div>'
            '<div class="meta">PeopleId %d &middot; %s &middot; DOB %s &middot; age %s now &middot; %s'
            ' &middot; <a href="/Person2/%d" target="_blank">open profile</a></div></div>'
            % (esc(t['Name']), t['PeopleId'], esc(t['Gender'] or '?'), esc(dob), age,
               esc(t['MemberStatus'] or ''), t['PeopleId']))


def render_audit(pid):
    t = fetch_target(pid)
    if not t:
        return '<div class="cap">No person found for PeopleId %d.</div>' % pid

    sessions = fetch_sessions(pid)
    co = fetch_co(pid, t['FamilyId'])

    # aggregates
    others_by_meeting = {}
    minors_by_meeting = {}
    for c in co:
        others_by_meeting[c['MeetingId']] = others_by_meeting.get(c['MeetingId'], 0) + 1
        if c['Minor']:
            minors_by_meeting[c['MeetingId']] = minors_by_meeting.get(c['MeetingId'], 0) + 1
    distinct_people = len(set(c['PeopleId'] for c in co))
    minor_rows = [c for c in co if c['Minor']]
    distinct_minors = len(set(c['PeopleId'] for c in minor_rows))

    h = []
    h.append(subject_block(t))
    h.append(window_form(pid))
    h.append('<a class="muted" href="?">&larr; audit a different person</a>')

    h.append('<div class="cards">')
    h.append('<div class="card"><div class="lbl">Sessions present</div><div class="val">%d</div>'
             '<div class="sub">%s</div></div>' % (len(sessions), esc(WINDOW_LABEL)))
    h.append('<div class="card"><div class="lbl">People co-present</div><div class="val">%d</div>'
             '<div class="sub">%d co-attendance records</div></div>' % (distinct_people, len(co)))
    h.append('<div class="card minor"><div class="lbl">Distinct minors</div><div class="val">%d</div>'
             '<div class="sub">under %d at the time</div></div>' % (distinct_minors, MINOR_AGE))
    h.append('<div class="card minor"><div class="lbl">Minor co-attendances</div><div class="val">%d</div></div>' % len(minor_rows))
    h.append('</div>')

    if len(co) >= CAP:
        h.append('<div class="cap">Row cap reached (%d). Narrow with the date window for a complete extract.</div>' % CAP)

    # ---- minors detail ----
    h.append('<div class="sec">Minors co-present <span class="c">(under %d at time of attendance)</span></div>' % MINOR_AGE)
    if not minor_rows:
        h.append('<div class="muted">No minors were co-present in this window.</div>')
    else:
        bykid = {}
        for c in minor_rows:
            bykid.setdefault(c['PeopleId'], []).append(c)
        order = sorted(bykid.keys(), key=lambda k: -len(bykid[k]))
        h.append('<div class="toolbar"><button class="btn alt" onclick="paaExport(\'paaMinors\',\'minors_detail.csv\')">&#8681; Export CSV</button>'
                 '<button class="btn" onclick="paaPrint()">Print / Save</button></div>')
        h.append('<table id="paaMinors"><thead><tr class="hdr">')
        for i, cn in enumerate(['Minor', 'Gender', 'DOB', 'Age range', 'Sessions', 'First', 'Last', 'Where', 'Family']):
            h.append('<th onclick="paaSort(\'paaMinors\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
        h.append('</tr></thead><tbody>')
        for kid in order:
            rs = bykid[kid]
            ages = sorted(set(c['Age'] for c in rs if c['Age'] is not None))
            arange = ('%d' % ages[0]) if len(ages) == 1 else ('%d-%d' % (ages[0], ages[-1])) if ages else '?'
            dates = sorted(set(c['Date'] for c in rs))
            orgs = ', '.join(sorted(set(c['Org'] for c in rs)))
            c0 = rs[0]
            fam = ' <span class="fam">SAME FAMILY</span>' if c0['SameFam'] else ''
            h.append('<tr data-pid="%d" data-oid="%d" data-minor="1">'
                     '<td><a href="?pid=%d&amp;co=%d">%s</a>%s'
                     ' <a href="/Person2/%d" target="_blank" class="muted" style="font-size:11px;" title="Profile">&#9432;</a></td>'
                     '<td>%s</td><td>%s</td><td class="num minorflag">%s</td><td class="num">%d</td>'
                     '<td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                     % (kid, c0['OrgId'], pid, kid, esc(c0['Name']), fam, kid,
                        esc(c0['Gender']), esc(c0['DOB']), arange, len(rs),
                        dates[0], dates[-1], esc(orgs), ('yes' if c0['SameFam'] else '')))
        h.append('</tbody></table>')

    # ---- sessions the target attended ----
    h.append('<div class="sec">Sessions present <span class="c">(%d meetings)</span></div>' % len(sessions))
    if sessions:
        sess_sorted = sorted(sessions, key=lambda s: s['Date'], reverse=True)
        h.append('<table id="paaSess"><thead><tr class="hdr">')
        for i, cn in enumerate(['Date', 'Day', 'Organization', 'Target role', 'Others', 'Minors']):
            h.append('<th onclick="paaSort(\'paaSess\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
        h.append('</tr></thead><tbody>')
        for s in sess_sorted:
            mn = minors_by_meeting.get(s['MeetingId'], 0)
            h.append('<tr data-pid="0" data-oid="%d"><td>%s</td><td>%s</td>'
                     '<td><a href="/Organization/%d" target="_blank">%s</a></td>'
                     '<td>%s</td><td class="num">%d</td><td class="num %s">%d</td></tr>'
                     % (s['OrgId'], s['Date'], weekday(s['Date']), s['OrgId'], esc(s['Org']),
                        esc(s['Role'] or ''), others_by_meeting.get(s['MeetingId'], 0),
                        'minorflag' if mn else '', mn))
        h.append('</tbody></table>')

    # ---- full co-attendance ----
    h.append('<div class="sec">Co-attendance detail <span class="c">(%d records, %d people)</span></div>'
             % (len(co), distinct_people))
    h.append('<div class="toolbar">'
             '<button class="btn alt" onclick="paaExport(\'paaCo\',\'coattendance.csv\')">&#8681; Export CSV</button>'
             '<button class="btn" onclick="paaPrint()">Print / Save</button>'
             '<label class="chk"><input type="checkbox" id="paaHideAdult" onchange="paaFilter(\'paaCo\')"> Minors only</label>'
             '<span class="count"><b id="paaShown">%d</b> shown</span></div>' % len(co))
    h.append('<table id="paaCo"><thead><tr class="hdr">')
    for i, cn in enumerate(['Date', 'Day', 'Organization', 'Person', 'Age', 'Minor', 'Role', 'Gender', 'DOB']):
        h.append('<th onclick="paaSort(\'paaCo\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
    h.append('</tr><tr class="flt">')
    for i in range(9):
        h.append('<th><input type="text" oninput="paaFilter(\'paaCo\')" placeholder="filter"></th>')
    h.append('</tr></thead><tbody>')
    for c in co:
        minorcell = '<span class="minorflag">MINOR</span>' if c['Minor'] else ''
        fam = ' <span class="fam">FAM</span>' if c['SameFam'] else ''
        agecell = ('%d' % c['Age']) if c['Age'] is not None else '?'
        h.append('<tr data-pid="%d" data-oid="%d" data-minor="%d">'
                 '<td>%s</td><td>%s</td>'
                 '<td><a href="/Organization/%d" target="_blank">%s</a></td>'
                 '<td><a href="?pid=%d&amp;co=%d">%s</a>%s'
                 ' <a href="/Person2/%d" target="_blank" class="muted" style="font-size:11px;">&#9432;</a></td>'
                 '<td class="num">%s</td><td class="num">%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                 % (c['PeopleId'], c['OrgId'], (1 if c['Minor'] else 0),
                    c['Date'], weekday(c['Date']), c['OrgId'], esc(c['Org']),
                    pid, c['PeopleId'], esc(c['Name']), fam, c['PeopleId'],
                    agecell, minorcell, esc(c['Role'] or ''), esc(c['Gender'] or ''), esc(c['DOB'] or '')))
    h.append('</tbody></table>')
    return ''.join(h)


def render_drill(pid, co_pid):
    t = fetch_target(pid)
    if not t:
        return '<div class="cap">No person found.</div>'
    co = [c for c in fetch_co(pid, t['FamilyId']) if c['PeopleId'] == co_pid]
    name = co[0]['Name'] if co else ('Person %d' % co_pid)
    h = [subject_block(t)]
    h.append('<a class="muted" href="?pid=%d">&larr; back to audit</a>' % pid)
    h.append('<div class="sec">Sessions <a href="/Person2/%d" target="_blank">%s</a> was co-present with the subject '
             '<span class="c">(%d)</span></div>' % (co_pid, esc(name), len(co)))
    if not co:
        h.append('<div class="muted">No co-attended sessions found in this window.</div>')
        return ''.join(h)
    h.append('<button class="btn" onclick="paaPrint()">Print / Save</button>')
    h.append('<table><thead><tr><th>Date</th><th>Day</th><th>Organization</th><th class="num">Their age</th><th>Their role</th></tr></thead><tbody>')
    for c in sorted(co, key=lambda x: x['Date']):
        agecell = ('%d' % c['Age']) if c['Age'] is not None else '?'
        cls = ' class="num minorflag"' if c['Minor'] else ' class="num"'
        h.append('<tr><td>%s</td><td>%s</td><td><a href="/Organization/%d" target="_blank">%s</a></td>'
                 '<td%s>%s</td><td>%s</td></tr>'
                 % (c['Date'], weekday(c['Date']), c['OrgId'], esc(c['Org']), cls, agecell, esc(c['Role'] or '')))
    h.append('</tbody></table>')
    return ''.join(h)


print CSS
print PRINT_JS
print '<div class="paa">'
print '<h2>Person Attendance Audit</h2>'
print ('<p class="lede">Point this at one person to reconstruct, from actual attendance, every session they were '
       'present for and <b>everyone else present at those same sessions</b> &mdash; with each person&#39;s age at '
       'the time and a minor flag. Authorized safe-church / incident review only. Read-only.</p>')
try:
    if PID and CO:
        print render_drill(PID, CO)
    elif PID:
        print render_audit(PID)
    else:
        print render_search()
except Exception as e:
    print '<div style="color:#b91c1c;font-family:monospace;white-space:pre-wrap;">Error: {0}\n{1}</div>'.format(str(e), traceback.format_exc())
print '</div>'
