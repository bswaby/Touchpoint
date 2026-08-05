#roles=Edit
# -*- coding: utf-8 -*-
#####################################################################
# TPxi_InProgressRegistrations  --  READ-ONLY report
#####################################################################
# People who STARTED a registration but have not finished it, across BOTH
# classic (old) and new registration forms, with a "Completed?" column so
# you can tell true abandoners from people who actually finished and left a
# stale in-progress record behind.
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
#
#####################################################################

import re
import traceback

model.Header = "In-Progress Registrations (read-only)"

CAP = 5000  # hard row cap so the page never tries to render a runaway set


def _param(name, default=''):
    try:
        v = getattr(model.Data, name)
        v = str(v).strip()
        return v if v else default
    except:
        return default


def _isodate(s):
    return s if re.match(r'^\d{4}-\d{2}-\d{2}$', s or '') else ''


DAYS = _param('days', '30')
try:
    DAYS = max(1, int(DAYS))
except:
    DAYS = 30
FROM = _isodate(_param('from', ''))
TO = _isodate(_param('to', ''))
# include classic sessions that never got tied to a person (anonymous starts)
INCLUDE_UNID = _param('unid', '') in ('1', 'true', 'on', 'yes')
UNID_SQL = '' if INCLUDE_UNID else 'AND p.PeopleId IS NOT NULL'

# what to show: in-progress (unfinished), completed (finished), or all
STATUS = _param('status', 'inprogress').lower()
if STATUS not in ('inprogress', 'completed', 'all'):
    STATUS = 'inprogress'
# per-mode: the status predicate for each source, and which date the window +
# the "Started" column key off (start date for in-progress, completion date
# for completed).
if STATUS == 'completed':
    C_STAT = "ISNULL(e.completed, 0) <> 0"
    C_DATE = "e.Stamp"
    N_STAT = "rp.CompletedDate IS NOT NULL"
    N_DATE = "rp.CompletedDate"
    DATE_HDR = "Completed"
elif STATUS == 'all':
    C_STAT = "ISNULL(e.abandoned, 0) = 0"
    C_DATE = "e.Stamp"
    N_STAT = "1 = 1"
    N_DATE = "COALESCE(rp.CompletedDate, r.CreatedDate)"
    DATE_HDR = "Date"
else:  # inprogress
    C_STAT = "ISNULL(e.completed, 0) = 0 AND ISNULL(e.abandoned, 0) = 0"
    C_DATE = "e.Stamp"
    N_STAT = "rp.CompletedDate IS NULL"
    N_DATE = "r.CreatedDate"
    DATE_HDR = "Started"

_verb = {'completed': 'completed', 'all': 'active', 'inprogress': 'started'}[STATUS]
if FROM or TO:
    WINDOW_LABEL = "%s %s to %s" % (_verb, FROM or 'start', TO or 'now')
else:
    WINDOW_LABEL = "%s in the last %d days" % (_verb, DAYS)


def date_pred(col):
    # SQL predicate bounding a date column to the chosen window. Inputs are
    # validated to YYYY-MM-DD, so no injection surface.
    if FROM and TO:
        return "%s >= '%s' AND %s < DATEADD(DAY, 1, '%s')" % (col, FROM, col, TO)
    if FROM:
        return "%s >= '%s'" % (col, FROM)
    if TO:
        return "%s < DATEADD(DAY, 1, '%s')" % (col, TO)
    return "%s >= DATEADD(DAY, -%d, GETDATE())" % (col, DAYS)


def rows_sql():
    return """
    SET NOCOUNT ON;
    SELECT TOP %d * FROM (

      SELECT 'Classic' AS FormType, p.PeopleId, ISNULL(p.Name, '(unidentified)') AS PersonName,
             e.OrganizationId, o.OrganizationName,
             CONVERT(varchar(19), %s, 126) AS Started,
             pr.Id AS ProgId, ISNULL(pr.Name, '(no program)') AS Program,
             ISNULL(c.Description, '') AS Campus,
             CASE WHEN EXISTS (SELECT 1 FROM dbo.RegistrationData rd
                               WHERE rd.OrganizationId = e.OrganizationId AND rd.UserPeopleId = e.UserPeopleId
                                 AND ISNULL(rd.completed, 0) <> 0)
                    OR EXISTS (SELECT 1 FROM dbo.Registration r2
                               WHERE r2.OrganizationId = e.OrganizationId AND r2.PeopleId = e.UserPeopleId
                                 AND r2.CompletedDate IS NOT NULL)
                    OR EXISTS (SELECT 1 FROM dbo.OrganizationMembers om
                               WHERE om.OrganizationId = e.OrganizationId AND om.PeopleId = e.UserPeopleId)
                  THEN 1 ELSE 0 END AS CompletedLater
      FROM dbo.RegistrationData e
      JOIN dbo.Organizations o ON o.OrganizationId = e.OrganizationId
      LEFT JOIN dbo.People p ON p.PeopleId = e.UserPeopleId
      LEFT JOIN dbo.Division d ON d.Id = o.DivisionId
      LEFT JOIN dbo.Program pr ON pr.Id = d.ProgId
      LEFT JOIN lookup.Campus c ON c.Id = o.CampusId
      WHERE %s
        AND ISNULL(o.RegistrationTypeId, 0) <> 26
        %s
        AND %s

      UNION ALL

      -- New forms: identity + progress live in RegPeople (one row per
      -- registrant). PeopleId may be NULL while in progress (not linked to a
      -- Person yet), so the name comes from RegPeople.FirstName/LastName.
      SELECT 'New' AS FormType, rp.PeopleId,
             LTRIM(RTRIM(ISNULL(rp.FirstName, '') + ' ' + ISNULL(rp.LastName, ''))) AS PersonName,
             r.OrganizationId, o.OrganizationName,
             CONVERT(varchar(19), %s, 126) AS Started,
             pr.Id AS ProgId, ISNULL(pr.Name, '(no program)') AS Program,
             ISNULL(rc.Description, ISNULL(c.Description, '')) AS Campus,
             CASE WHEN (rp.PeopleId IS NOT NULL AND (
                          EXISTS (SELECT 1 FROM dbo.OrganizationMembers om
                                  WHERE om.OrganizationId = r.OrganizationId AND om.PeopleId = rp.PeopleId)
                       OR EXISTS (SELECT 1 FROM dbo.RegistrationData rd
                                  WHERE rd.OrganizationId = r.OrganizationId AND rd.UserPeopleId = rp.PeopleId
                                    AND ISNULL(rd.completed, 0) <> 0)))
                    OR EXISTS (SELECT 1 FROM dbo.RegPeople rp2
                               JOIN dbo.Registration r2 ON r2.RegistrationId = rp2.RegistrationId
                               WHERE r2.OrganizationId = r.OrganizationId AND rp2.CompletedDate IS NOT NULL
                                 AND ((rp.PeopleId IS NOT NULL AND rp2.PeopleId = rp.PeopleId)
                                   OR (NULLIF(rp.Email, '') IS NOT NULL AND rp2.Email = rp.Email)))
                  THEN 1 ELSE 0 END AS CompletedLater
      FROM dbo.RegPeople rp
      JOIN dbo.Registration r ON r.RegistrationId = rp.RegistrationId
      JOIN dbo.Organizations o ON o.OrganizationId = r.OrganizationId
      LEFT JOIN dbo.Division d ON d.Id = o.DivisionId
      LEFT JOIN dbo.Program pr ON pr.Id = d.ProgId
      LEFT JOIN lookup.Campus c ON c.Id = o.CampusId
      LEFT JOIN lookup.Campus rc ON rc.Id = rp.CampusId
      WHERE %s
        AND %s

    ) z
    ORDER BY z.Started DESC;
    """ % (CAP, C_DATE, C_STAT, UNID_SQL, date_pred(C_DATE), N_DATE, N_STAT, date_pred(N_DATE))


def fetch():
    rows = []
    for r in q.QuerySql(rows_sql()):
        pid = int(r.PeopleId) if r.PeopleId else 0
        name = _u(r.PersonName).strip()
        if not name:
            name = ('Person ' + str(pid)) if pid else '(unnamed)'
        rows.append({
            'FormType': r.FormType,
            'PeopleId': pid,
            'Person': name,
            'OrgId': r.OrganizationId,
            'Org': r.OrganizationName or ('Org ' + str(r.OrganizationId)),
            'Started': (r.Started or '')[:10],
            'ProgId': int(r.ProgId) if r.ProgId else 0,
            'Program': r.Program or '(no program)',
            'Campus': r.Campus or '',
            'Done': (int(r.CompletedLater or 0) == 1),
        })
    return rows


CSS = """
<style>
  .ipr { font-family:'Segoe UI',-apple-system,Arial,sans-serif; color:#1e293b; max-width:1180px; line-height:1.5; }
  .ipr h2 { color:#1f4e79; margin:0 0 2px; font-size:24px; }
  .ipr .lede { color:#64748b; font-size:13px; margin:0 0 14px; max-width:920px; }
  .ipr code { background:#eef2f7; padding:1px 5px; border-radius:3px; font-size:12px; }
  .ipr .cards { display:flex; gap:12px; flex-wrap:wrap; margin:0 0 16px; }
  .ipr .card { flex:1; min-width:150px; border:1px solid #e2e8f0; border-top:3px solid #1f4e79; border-radius:10px; padding:12px 14px; background:#fff; }
  .ipr .card .lbl { font-size:11px; text-transform:uppercase; letter-spacing:.6px; color:#64748b; font-weight:600; }
  .ipr .card .val { font-size:26px; font-weight:800; margin-top:3px; color:#1f4e79; }
  .ipr .card .sub { font-size:11px; color:#94a3b8; }
  .ipr .card.warn { border-top-color:#b45309; } .ipr .card.warn .val { color:#b45309; }
  .ipr .filters { display:flex; flex-wrap:wrap; align-items:flex-end; gap:14px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px 16px; margin:0 0 14px; }
  .ipr .filters label { font-size:11px; text-transform:uppercase; letter-spacing:.5px; color:#64748b; font-weight:600; display:flex; flex-direction:column; gap:3px; }
  .ipr .filters input { font-size:14px; padding:5px 8px; border:1px solid #cbd5e1; border-radius:6px; }
  .ipr .filters .go { background:#1f4e79; color:#fff; border:0; border-radius:6px; padding:7px 16px; font-size:14px; cursor:pointer; }
  .ipr .filters .reset { color:#64748b; font-size:12px; align-self:center; }
  .ipr .toolbar { display:flex; align-items:center; gap:14px; flex-wrap:wrap; margin:0 0 8px; }
  .ipr .btn { display:inline-block; background:#1f4e79; color:#fff; padding:7px 14px; border-radius:6px; font-size:13px; border:0; cursor:pointer; }
  .ipr .btn.alt { background:#0f766e; }
  .ipr .toolbar .chk { display:flex; align-items:center; gap:6px; font-size:13px; color:#334155; }
  .ipr .toolbar .count { color:#64748b; font-size:13px; margin-left:auto; }
  .ipr table { border-collapse:collapse; width:100%; font-size:13px; margin:2px 0 6px; }
  .ipr th { background:#eef2f7; color:#334155; text-align:left; padding:5px 9px; font-size:11px; text-transform:uppercase; letter-spacing:.4px; cursor:pointer; white-space:nowrap; }
  .ipr th:hover { background:#e2e8f0; }
  .ipr th .arrow { color:#94a3b8; font-weight:400; }
  .ipr thead .flt th { background:#fff; padding:3px 6px; cursor:auto; }
  .ipr thead .flt input { width:100%; box-sizing:border-box; font-size:12px; padding:3px 5px; border:1px solid #cbd5e1; border-radius:4px; font-weight:400; text-transform:none; letter-spacing:0; }
  .ipr td { padding:5px 9px; border-bottom:1px solid #f1f5f9; }
  .ipr a { color:#1f4e79; text-decoration:none; } .ipr a:hover { text-decoration:underline; }
  .ipr .pill { display:inline-block; border-radius:20px; padding:1px 9px; font-size:11px; font-weight:700; }
  .ipr .pill.new { background:#ede9fe; color:#6d28d9; } .ipr .pill.classic { background:#e0f2fe; color:#0369a1; }
  .ipr .done { color:#166534; font-weight:700; } .ipr .notdone { color:#b45309; font-weight:700; }
  .ipr .muted { color:#94a3b8; font-size:12px; }
  .ipr .cap { background:#fff7ed; border:1px solid #fed7aa; border-left:4px solid #ea580c; color:#7a4a00; border-radius:8px; padding:8px 14px; margin:0 0 12px; font-size:12.5px; }
</style>
"""


def _u(val):
    # Robustly coerce any value (Python unicode/str, .NET System.String, bytes)
    # to a unicode string. Names/org titles can carry accents or smart quotes.
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
    # ASCII-safe HTML escape: escapes &<>"' and emits numeric character
    # references (&#NNNN;) for every non-ASCII codepoint. The printed bytes are
    # pure ASCII (so IronPython's print never hits the 'ascii' codec at the .NET
    # boundary), while the browser and the UTF-8 CSV still show the real accent.
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


def build_js():
    # per-column filter + click-to-sort + hide-done + CSV export
    return """
<script>
function iprFilter(){
  var t=document.getElementById('iprTable'); if(!t) return;
  var ins=t.querySelectorAll('thead tr.flt input');
  var vals=[]; for(var i=0;i<ins.length;i++) vals.push(ins[i].value.toLowerCase());
  var hide=document.getElementById('iprHideDone').checked;
  var rows=t.querySelectorAll('tbody tr'); var shown=0;
  for(var r=0;r<rows.length;r++){
    var tds=rows[r].getElementsByTagName('td'); var ok=true;
    for(var c=0;c<vals.length;c++){ if(vals[c] && (tds[c].textContent||tds[c].innerText||'').toLowerCase().indexOf(vals[c])<0){ok=false;break;} }
    if(ok && hide && rows[r].getAttribute('data-done')==='1') ok=false;
    rows[r].style.display = ok?'':'none'; if(ok) shown++;
  }
  var s=document.getElementById('iprShown'); if(s) s.textContent=shown;
}
var iprDir={};
function iprSort(c){
  var t=document.getElementById('iprTable'); var tb=t.tBodies[0];
  var rows=[].slice.call(tb.rows);
  var dir=iprDir[c]=-(iprDir[c]||1);
  rows.sort(function(a,b){
    var x=(a.cells[c].textContent||'').trim().toLowerCase(), y=(b.cells[c].textContent||'').trim().toLowerCase();
    if(x<y) return -1*dir; if(x>y) return 1*dir; return 0;
  });
  for(var i=0;i<rows.length;i++) tb.appendChild(rows[i]);
  var ths=t.tHead.querySelectorAll('tr.hdr th');
  for(var j=0;j<ths.length;j++){ var a=ths[j].querySelector('.arrow'); if(a) a.textContent = (j===c) ? (dir>0?'\\u25B2':'\\u25BC') : '\\u21C5'; }
}
function iprCsvCell(v){ v=(v==null?'':(''+v)); return '"'+v.replace(/"/g,'""')+'"'; }
function iprExport(){
  var t=document.getElementById('iprTable'); if(!t) return;
  var head=['Program','Campus','Organization','OrgId','Person','PeopleId','Started','Form','Completed'];
  var lines=[head.map(iprCsvCell).join(',')];
  var rows=t.tBodies[0].rows;
  for(var r=0;r<rows.length;r++){ if(rows[r].style.display==='none') continue;
    var td=rows[r].cells;
    function ct(i){ return (td[i].textContent||'').trim(); }
    lines.push([ct(0),ct(1),ct(2),rows[r].getAttribute('data-oid'),ct(3),rows[r].getAttribute('data-pid'),ct(4),ct(5),ct(6)].map(iprCsvCell).join(','));
  }
  var csv=lines.join('\\r\\n');
  var blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='in-progress-registrations.csv'; document.body.appendChild(a); a.click();
  setTimeout(function(){ document.body.removeChild(a); URL.revokeObjectURL(a.href); },100);
}
function iprPrint(){
  var node=document.querySelector('.ipr'); if(!node){window.print();return;}
  var clone=node.cloneNode(true);
  var kill=clone.querySelectorAll('.btn,.filters,.toolbar .chk,thead tr.flt'); for(var i=0;i<kill.length;i++){kill[i].parentNode.removeChild(kill[i]);}
  // drop hidden (filtered-out) rows from the printout
  var rows=clone.querySelectorAll('tbody tr'); for(var r=0;r<rows.length;r++){ if(rows[r].style.display==='none') rows[r].parentNode.removeChild(rows[r]); }
  var css="*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}body{font-family:'Segoe UI',Arial,sans-serif;color:#1e293b;padding:20px}h2{color:#1f4e79}table{border-collapse:collapse;width:100%;font-size:11px}th{background:#eef2f7;text-align:left;padding:4px 8px}td{padding:4px 8px;border-bottom:1px solid #f1f5f9}.pill{border-radius:20px;padding:1px 8px;font-size:10px;font-weight:700}.pill.new{background:#ede9fe;color:#6d28d9}.pill.classic{background:#e0f2fe;color:#0369a1}.done{color:#166534;font-weight:700}.notdone{color:#b45309;font-weight:700}.cards{display:flex;gap:10px;margin-bottom:12px}.card{border:1px solid #e2e8f0;border-top:3px solid #1f4e79;border-radius:8px;padding:8px 12px}.card .val{font-size:20px;font-weight:800;color:#1f4e79}.card .lbl{font-size:10px;text-transform:uppercase;color:#64748b}a{color:#1e293b;text-decoration:none}.arrow{display:none}";
  var pw=window.open('','_blank'); if(!pw){alert('Allow popups to print');return;}
  pw.document.write('<!DOCTYPE html><html><head><title>In-Progress Registrations</title><style>'+css+'</style></head><body>'+clone.innerHTML+'</body></html>');
  pw.document.close(); pw.focus(); setTimeout(function(){pw.print();},350);
}
</script>
"""


def render(rows):
    total = len(rows)
    classic_n = sum(1 for r in rows if r['FormType'] == 'Classic')
    new_n = sum(1 for r in rows if r['FormType'] == 'New')
    done_n = sum(1 for r in rows if r['Done'])

    html = []

    # window bar (server-side bookends)
    st_opts = ''.join('<option value="%s"%s>%s</option>' % (v, ' selected' if STATUS == v else '', lbl)
                      for v, lbl in [('inprogress', 'In progress (not finished)'),
                                     ('completed', 'Completed'), ('all', 'All')])
    html.append('<form method="get" class="filters">'
                '<label>Show<select name="status">%s</select></label>'
                '<label>Last N days<input type="number" name="days" min="1" value="%d"></label>'
                '<label>From (optional)<input type="date" name="from" value="%s"></label>'
                '<label>To (optional)<input type="date" name="to" value="%s"></label>'
                '<label class="chk" style="flex-direction:row;align-items:center;gap:6px;text-transform:none;letter-spacing:0;font-size:13px;color:#334155;">'
                '<input type="checkbox" name="unid" value="1"%s> Include unidentified (anonymous starts)</label>'
                '<button type="submit" class="go">Apply</button>'
                '<a class="reset" href="?">reset</a>'
                '</form>' % (st_opts, DAYS, FROM, TO, ' checked' if INCLUDE_UNID else ''))

    total_lbl = {'inprogress': 'In progress', 'completed': 'Completed', 'all': 'Records'}[STATUS]
    html.append('<div class="cards">')
    html.append('<div class="card"><div class="lbl">%s</div><div class="val">%d</div>'
                '<div class="sub">%s</div></div>' % (total_lbl, total, esc(WINDOW_LABEL)))
    if STATUS != 'completed':
        html.append('<div class="card" style="border-top-color:#0f766e;"><div class="lbl">Not yet (actionable)</div>'
                    '<div class="val" style="color:#0f766e;">%d</div><div class="sub">still unfinished</div></div>' % (total - done_n))
        html.append('<div class="card warn"><div class="lbl">Already completed</div><div class="val">%d</div>'
                    '<div class="sub">stale; they finished</div></div>' % done_n)
    html.append('<div class="card"><div class="lbl">Classic (old)</div><div class="val">%d</div></div>' % classic_n)
    html.append('<div class="card"><div class="lbl">New forms</div><div class="val">%d</div></div>' % new_n)
    html.append('</div>')

    if total >= CAP:
        html.append('<div class="cap">Showing the most recent <b>%d</b> rows (cap reached). '
                    'Narrow the window with <b>Last N days</b> or a From/To range to see everything.</div>' % CAP)

    # toolbar: export / print / hide-done / live count
    html.append('<div class="toolbar">'
                '<button class="btn alt" onclick="iprExport()">&#8681; Export CSV</button>'
                '<button class="btn" onclick="iprPrint()">Print / Save</button>'
                '<label class="chk"><input type="checkbox" id="iprHideDone" onchange="iprFilter()"> Hide already-completed</label>'
                '<span class="count"><b id="iprShown">%d</b> shown</span>'
                '</div>' % total)

    if not rows:
        html.append('<div class="card" style="border-top-color:#94a3b8;">No in-progress registrations in this window.</div>')
        return ''.join(html)

    # single flat, sortable, column-filterable table
    cols = ['Program', 'Campus', 'Organization', 'Person', DATE_HDR, 'Form', 'Completed?']
    html.append('<table id="iprTable"><thead><tr class="hdr">')
    for i, cname in enumerate(cols):
        html.append('<th onclick="iprSort(%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cname))
    html.append('</tr><tr class="flt">')
    for i in range(len(cols)):
        html.append('<th><input type="text" oninput="iprFilter()" placeholder="filter"></th>')
    html.append('</tr></thead><tbody>')

    for r in rows:
        pill = '<span class="pill %s">%s</span>' % (r['FormType'].lower(), r['FormType'])
        done = ('<span class="done">completed</span>' if r['Done']
                else '<span class="notdone">not yet</span>')
        # new-form in-progress registrants may have no PeopleId yet -> no link
        person = (('<a href="/Person2/%d" target="_blank">%s</a>' % (r['PeopleId'], esc(r['Person'])))
                  if r['PeopleId'] else esc(r['Person']))
        html.append('<tr data-pid="%s" data-oid="%d" data-done="%d">'
                    '<td>%s</td><td>%s</td>'
                    '<td><a href="/Organization/%d" target="_blank">%s</a></td>'
                    '<td>%s</td>'
                    '<td>%s</td><td>%s</td><td>%s</td></tr>'
                    % ((str(r['PeopleId']) if r['PeopleId'] else ''), r['OrgId'], (1 if r['Done'] else 0),
                       esc(r['Program']), esc(r['Campus']) or '&mdash;',
                       r['OrgId'], esc(r['Org']), person,
                       r['Started'] or '&mdash;', pill, done))
    html.append('</tbody></table>')
    return ''.join(html)


print CSS
print build_js()
print '<div class="ipr">'
print '<h2>In-Progress Registrations</h2>'
print ('<p class="lede">People who started a registration but have not finished, across <b>both classic '
       'and new</b> forms, within the chosen window. The <b>Completed?</b> column flags stale records &mdash; '
       'people who actually finished (a completed registration or an enrollment for that org exists). Click a '
       'column to sort, type under any header to filter, and use <b>Export CSV</b> for the (filtered) rows. '
       'Read-only.</p>')
try:
    print render(fetch())
except Exception as e:
    print '<div style="color:#b91c1c;font-family:monospace;white-space:pre-wrap;">Error: {0}\n{1}</div>'.format(str(e), traceback.format_exc())
print '</div>'
