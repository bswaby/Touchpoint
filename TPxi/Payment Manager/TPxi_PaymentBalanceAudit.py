#####################################################################
# TPxi_PaymentBalanceAudit  --  READ-ONLY validation harness
#####################################################################
# What this checks
#   Payment Manager sources balances from dbo.TransactionSummary, which
#   matches how TouchPoint computes "Due" for CLASSIC registrations
#   (verified against bvcms OrgMemberModel.cs).
#
#   For NEW-FORM registrations (RegistrationTypeId = 26 =
#   RegistrationTypeCode.RegistrationForm) TouchPoint computes Due from
#   RegistrationTransactionSummary:  MAX(0, TotalAmt - TotalPaid).
#   Payment Manager still reads TransactionSummary there, so it can disagree
#   with the TouchPoint UI on type-26 orgs only. This report lists every
#   type-26 org/person where the tool ("Tool") differs from TouchPoint ("TP")
#   so each can be confirmed against the enrollment "Due" line before any
#   code change. Classic orgs already match TP and are not shown.
#
#   Read-only (SELECTs only). One per-person query (restricted to type-26
#   orgs for performance); all totals/grouping happen in Python -- q.QuerySql
#   would not run the full-table SUM aggregate over these views.
#
# Usage
#   /PyScript/TPxi_PaymentBalanceAudit           -> type-26 orgs that differ
#   /PyScript/TPxi_PaymentBalanceAudit?org=1234  -> per-person drill-down
#####################################################################

import traceback

model.Header = "Payment Balance Audit -- new-form (type 26) orgs (read-only)"


# One per-person balance query. Subqueries are restricted to type-26 orgs so
# the TransactionSummary / RegistrationTransactionSummary views are only
# computed for those ~200 orgs, not the whole database. org_filter optionally
# adds "AND o.OrganizationId = N" for the drill-down.
def rows_sql(org_filter=""):
    return """
    SELECT om.OrganizationId AS OrgId, o.OrganizationName, om.PeopleId, p.Name2,
        CASE WHEN ISNULL(ts.D,0) - ISNULL(sup.S,0) > 0
             THEN ISNULL(ts.D,0) - ISNULL(sup.S,0) ELSE 0 END AS OldBal,
        CASE WHEN ISNULL(rtp.Amt,0) - ISNULL(rtp.Paid,0) > 0
             THEN ISNULL(rtp.Amt,0) - ISNULL(rtp.Paid,0) ELSE 0 END AS TpBal
    FROM dbo.OrganizationMembers om
    JOIN dbo.Organizations o
      ON o.OrganizationId = om.OrganizationId AND o.RegistrationTypeId = 26 %s
    LEFT JOIN dbo.People p ON p.PeopleId = om.PeopleId
    LEFT JOIN (
        SELECT t.OrganizationId, t.PeopleId, SUM(t.IndDue) AS D
        FROM dbo.TransactionSummary t
        JOIN dbo.Organizations o2 ON o2.OrganizationId = t.OrganizationId AND o2.RegistrationTypeId = 26
        WHERE t.IsLatestTransaction = 1
        GROUP BY t.OrganizationId, t.PeopleId
    ) ts ON ts.OrganizationId = om.OrganizationId AND ts.PeopleId = om.PeopleId
    LEFT JOIN (
        SELECT r.OrganizationId, r.PeopleId, SUM(r.TotalAmt) AS Amt, SUM(r.TotalPaid) AS Paid
        FROM dbo.RegistrationTransactionSummary r
        JOIN dbo.Organizations o3 ON o3.OrganizationId = r.OrganizationId AND o3.RegistrationTypeId = 26
        WHERE r.IsLatestTransaction = 1
        GROUP BY r.OrganizationId, r.PeopleId
    ) rtp ON rtp.OrganizationId = om.OrganizationId AND rtp.PeopleId = om.PeopleId
    LEFT JOIN (
        SELECT GoerId, OrgId, SUM(Amount) AS S
        FROM dbo.GoerSenderAmounts
        WHERE ISNULL(InActive, 0) = 0 AND SupporterId <> GoerId
        GROUP BY GoerId, OrgId
    ) sup ON sup.GoerId = om.PeopleId AND sup.OrgId = om.OrganizationId
    WHERE (ISNULL(ts.D,0) - ISNULL(sup.S,0) > 0.01
        OR ISNULL(rtp.Amt,0) - ISNULL(rtp.Paid,0) > 0.01)
    """ % (org_filter,)


def fetch(org_filter=""):
    out = []
    for r in q.QuerySql(rows_sql(org_filter)):
        out.append({
            'OrgId': r.OrgId,
            'OrgName': r.OrganizationName or ('Org ' + str(r.OrgId)),
            'PeopleId': r.PeopleId,
            'Name2': r.Name2 or ('Person ' + str(r.PeopleId)),
            'Old': float(r.OldBal or 0),
            'Tp': float(r.TpBal or 0),
        })
    return out


def money(v):
    try:
        return "${:,.2f}".format(float(v or 0))
    except:
        return "$0.00"


CSS = """
<style>
  .pba { font-family: Segoe UI, Arial, sans-serif; color:#333; max-width:1050px; }
  .pba h2 { color:#1f4e79; margin:0 0 4px; }
  .pba .lede { color:#666; font-size:13px; margin:0 0 14px; line-height:1.55; }
  .pba code { background:#f3f3f3; padding:1px 5px; border-radius:3px; }
  .pba .cards { display:flex; gap:12px; flex-wrap:wrap; margin-bottom:18px; }
  .pba .card { flex:1; min-width:170px; border:1px solid #e3e3e3; border-radius:8px; padding:12px 14px; background:#fff; }
  .pba .card .lbl { font-size:11px; text-transform:uppercase; letter-spacing:.5px; color:#888; }
  .pba .card .val { font-size:22px; font-weight:700; margin-top:2px; }
  .pba .card .sub { font-size:11px; color:#999; margin-top:2px; }
  .pba .miss .val { color:#c0392b; }
  .pba .over .val { color:#2980b9; }
  .pba table { border-collapse:collapse; width:100%; font-size:13px; }
  .pba th { background:#1f4e79; color:#fff; text-align:left; padding:7px 9px; font-weight:600; }
  .pba td { padding:6px 9px; border-bottom:1px solid #eee; }
  .pba tr:hover td { background:#f5f9ff; }
  .pba .num { text-align:right; font-variant-numeric:tabular-nums; }
  .pba .delta-pos { color:#c0392b; font-weight:600; }
  .pba .delta-neg { color:#2980b9; font-weight:600; }
  .pba a { color:#1f4e79; text-decoration:none; }
  .pba a:hover { text-decoration:underline; }
  .pba .back { display:inline-block; margin-bottom:12px; font-size:13px; }
  .pba .warn { background:#fff8e1; border:1px solid #ffe0a3; border-radius:6px; padding:8px 12px; font-size:12px; color:#7a5b00; margin-bottom:16px; }
  .pba .ok { background:#eafaf0; border:1px solid #b6e6c8; border-radius:6px; padding:10px 14px; font-size:13px; color:#1e6b3a; }
</style>
"""


def render_summary(rows):
    miss_n = miss_d = over_n = over_d = diff_n = 0
    orgs = {}
    for x in rows:
        old, tp = x['Old'], x['Tp']
        differs = abs(old - tp) > 0.01
        if old < 0.01 and tp > 0.01:
            miss_n += 1; miss_d += tp
        elif tp < 0.01 and old > 0.01:
            over_n += 1; over_d += old
        elif old > 0.01 and tp > 0.01 and differs:
            diff_n += 1
        o = orgs.get(x['OrgId'])
        if o is None:
            o = {'name': x['OrgName'], 'old': 0.0, 'tp': 0.0, 'differ': 0}
            orgs[x['OrgId']] = o
        o['old'] += old
        o['tp'] += tp
        if differs:
            o['differ'] += 1

    html = '<div class="cards">'
    html += '<div class="card miss"><div class="lbl">Tool shows $0, TP owes</div><div class="val">{0}</div><div class="sub">{1} people</div></div>'.format(money(miss_d), miss_n)
    html += '<div class="card over"><div class="lbl">Tool owes, TP shows $0</div><div class="val">{0}</div><div class="sub">{1} people</div></div>'.format(money(over_d), over_n)
    html += '<div class="card"><div class="lbl">Both owe, amounts differ</div><div class="val">{0}</div><div class="sub">people</div></div>'.format(diff_n)
    html += '</div>'

    flagged = [(oid, o) for oid, o in orgs.items() if o['differ'] > 0]
    flagged.sort(key=lambda t: abs(t[1]['tp'] - t[1]['old']), reverse=True)
    if not flagged:
        return html + ('<div class="ok"><b>No differences found.</b> On every new-form (type 26) org, Payment '
                       'Manager\'s balance matches TouchPoint. Nothing to fix.</div>')
    html += '<table><thead><tr><th>Organization (type 26)</th><th class="num">Tool (current)</th><th class="num">TouchPoint</th><th class="num">Delta</th><th class="num">People differ</th></tr></thead><tbody>'
    for oid, o in flagged:
        delta = o['tp'] - o['old']
        cls = 'delta-pos' if delta > 0 else 'delta-neg'
        sign = '+' if delta > 0 else ''
        html += '<tr><td><a href="?org={0}">{1}</a></td>'.format(oid, o['name'])
        html += '<td class="num">{0}</td><td class="num">{1}</td>'.format(money(o['old']), money(o['tp']))
        html += '<td class="num {0}">{1}{2}</td><td class="num">{3}</td></tr>'.format(cls, sign, money(delta), o['differ'])
    html += '</tbody></table>'
    return html


def render_drill(org_id):
    org_i = int(org_id)
    rows = fetch("AND o.OrganizationId = " + str(org_i))
    rows = [x for x in rows if x['Old'] > 0.01 or x['Tp'] > 0.01]
    rows.sort(key=lambda x: (abs(x['Tp'] - x['Old']), x['Tp']), reverse=True)
    oname = rows[0]['OrgName'] if rows else ('Org ' + str(org_i))

    html = '<a class="back" href="?">&larr; All new-form orgs</a>'
    html += '<h3>{0} <span style="color:#999;font-weight:400;">(org {1}, type 26)</span></h3>'.format(oname, org_i)
    html += '<table><thead><tr><th>Person</th><th class="num">Tool (current)</th><th class="num">TouchPoint</th><th class="num">Delta</th></tr></thead><tbody>'
    for x in rows:
        delta = x['Tp'] - x['Old']
        if abs(delta) < 0.01:
            cls, sign = '', ''
        elif delta > 0:
            cls, sign = 'delta-pos', '+'
        else:
            cls, sign = 'delta-neg', ''
        html += '<tr><td><a href="/Person2/{0}#enrollment" target="_blank">{1}</a></td>'.format(x['PeopleId'], x['Name2'])
        html += '<td class="num">{0}</td><td class="num">{1}</td>'.format(money(x['Old']), money(x['Tp']))
        html += '<td class="num {0}">{1}{2}</td></tr>'.format(cls, sign, money(delta))
    html += '</tbody></table>'
    if not rows:
        html += '<p>No one in this org has a balance in either source.</p>'
    return html


# ------------------------------------------------------------------ #
org_param = None
try:
    if hasattr(model.Data, 'org') and str(model.Data.org).strip():
        org_param = str(model.Data.org).strip()
except:
    org_param = None

print CSS
print '<div class="pba">'
print '<h2>Payment Balance Audit</h2>'
print ('<p class="lede"><b>Tool (current)</b> = what Payment Manager reports, from <code>TransactionSummary</code>. '
       '<b>TouchPoint</b> = how TP itself computes Due for <b>new-form (RegistrationTypeId&nbsp;=&nbsp;26)</b> orgs, from '
       '<code>RegistrationTransactionSummary</code> (verified against bvcms <code>OrgMemberModel.cs</code>). '
       'Classic orgs already match TP and are not listed. Both sides net supporters and clamp at $0. '
       '<b style="color:#c0392b;">Positive delta = the tool misses a balance the TP UI shows.</b> Read-only.</p>')
print ('<div class="warn">Confirm each flagged person against the enrollment dialog\'s '
       '<b>Transaction Amounts &rarr; Due</b> line in TouchPoint before changing any code.</div>')

try:
    if org_param:
        print render_drill(org_param)
    else:
        print render_summary(fetch())
except Exception as e:
    print '<div style="color:#c0392b;font-family:monospace;white-space:pre-wrap;">Error: {0}\n{1}</div>'.format(str(e), traceback.format_exc())

print '</div>'
