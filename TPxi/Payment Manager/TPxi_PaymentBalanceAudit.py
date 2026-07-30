#####################################################################
# TPxi_PaymentBalanceAudit  --  READ-ONLY surface-consistency audit
#####################################################################
# PURPOSE
#   Surface every registration where TouchPoint's OWN balance
#   calculations disagree with each other. This is evidence for
#   TouchPoint support, not a collections list. It does not care who
#   owes; it cares where the surfaces contradict.
#
# THE FOUR SURFACES (each a different TouchPoint calculation)
#   Ledger   Transaction.amtdue on the latest transaction in the chain.
#            = the transactions screen "Amount Due" and the Trip Balance.
#   PayLink  SUM(TransactionSummary.IndDue) WHERE RegId = member.TranId.
#            = CmsData PayInfo, the {amountdue} replacement, the PayLink.
#   SB       TransactionSummary.IndDue of the highest-RegId row.
#            = the "Has Balance in Involvement" Search Builder condition
#              (Condition.Enrollments.cs HasBalance / HasBalanceInCurrentOrg).
#   Modal    Involvement pop-up "Due" line (CmsWeb OrgMemberModel):
#              new-form (RegistrationTypeId 26): MAX(0, RegistrationTransaction
#                Summary.TotalAmt - TotalPaid), IsLatestTransaction = 1.
#              classic: MAX(0, ledger - mission-trip supporter gifts),
#                matching MAX(0, fee - dbo.TotalPaid).
#
# WHAT THE DISAGREEMENTS MEAN
#   Mission trip   Supporter gifts (GoerSenderAmounts) and fee adjustments
#                  post differently to each surface, so the ledger can go
#                  negative while PayLink/SB show a stale positive and the
#                  modal nets to $0. All four disagree.
#   New-form sync  RegistrationTransactionSummary (the modal source)
#                  disagrees with TransactionSummary (PayLink/SB) and the
#                  transaction ledger.
#   PayLink stale  PayInfo sums IndDue on the member's ORIGINAL TranId,
#                  which is stale after later transactions; the ledger and
#                  SB read the latest.
#
#   Read-only (SELECTs only). Verified against the live transactions
#   screen and involvement dialog.
#
# USAGE
#   /PyScript/TPxi_PaymentBalanceAudit           -> all inconsistencies
#   /PyScript/TPxi_PaymentBalanceAudit?org=1234  -> one involvement
#####################################################################

import traceback

model.Header = "Registration Balance Consistency Audit (read-only)"


# All four surfaces per member. Scalar subqueries only (q.QuerySql-safe).
# Filtered in an outer wrapper to just the rows where they disagree.
def rows_sql(org_filter=""):
    # Level 1: raw surface values (Ledger/PayLink/SB can be negative = overpaid/credit).
    inner = """
        SELECT om.OrganizationId AS OrgId, o.OrganizationName, o.RegistrationTypeId AS RegType,
            ISNULL(o.IsMissionTrip, 0) AS Trip, om.PeopleId, om.TranId AS TranId, p.Name2,
            ISNULL((SELECT TOP 1 ts.TotDue FROM dbo.TransactionSummary ts
                    WHERE ts.PeopleId = om.PeopleId AND ts.OrganizationId = om.OrganizationId
                      AND ts.IsLatestTransaction = 1
                    ORDER BY ts.RegId DESC), 0) AS LedgerRaw,
            -- PayLink = the member-dialog "PayLink" button (StandardPayLink, lt=1), which skips the
            -- new-form redirect and always runs PaymentForm.AmountDueTrans (first TotDue>0 row).
            CASE WHEN o.IsMissionTrip = 1
                 THEN ISNULL((SELECT TOP 1 CASE WHEN (tt.IndAmt - ((ISNULL((SELECT TOP 1 tp1.IndPaid FROM dbo.TransactionSummary tp1 WHERE tp1.RegId = om.TranId AND tp1.PeopleId = om.PeopleId ORDER BY tp1.TranDate DESC), 0) + ISNULL((SELECT SUM(g2.Amount) FROM dbo.GoerSenderAmounts g2 WHERE g2.GoerId = om.PeopleId AND g2.OrgId = om.OrganizationId AND g2.SupporterId <> om.PeopleId AND g2.Created > om.EnrollmentDate), 0)) - ISNULL(tt.Donation, 0))) > 0
                                                THEN (tt.IndAmt - ((ISNULL((SELECT TOP 1 tp1.IndPaid FROM dbo.TransactionSummary tp1 WHERE tp1.RegId = om.TranId AND tp1.PeopleId = om.PeopleId ORDER BY tp1.TranDate DESC), 0) + ISNULL((SELECT SUM(g2.Amount) FROM dbo.GoerSenderAmounts g2 WHERE g2.GoerId = om.PeopleId AND g2.OrgId = om.OrganizationId AND g2.SupporterId <> om.PeopleId AND g2.Created > om.EnrollmentDate), 0)) - ISNULL(tt.Donation, 0))) ELSE 0 END
                              FROM dbo.TransactionSummary tt
                              WHERE tt.RegId = (SELECT TOP 1 OriginalId FROM dbo.[Transaction] WHERE Id = om.TranId)
                                AND tt.TotDue > 0 AND tt.OrganizationId = om.OrganizationId), 0)
                 ELSE ISNULL((SELECT TOP 1 tt.TotDue FROM dbo.TransactionSummary tt
                              WHERE tt.RegId = (SELECT TOP 1 OriginalId FROM dbo.[Transaction] WHERE Id = om.TranId)
                                AND tt.TotDue > 0 AND tt.OrganizationId = om.OrganizationId), 0)
            END AS PayLinkRaw,
            ISNULL((SELECT TOP 1 ts.IndDue FROM dbo.TransactionSummary ts
                    WHERE ts.PeopleId = om.PeopleId AND ts.OrganizationId = om.OrganizationId
                    ORDER BY ts.RegId DESC), 0) AS SBRaw,
            CASE WHEN o.RegistrationTypeId = 26
                 THEN ISNULL((SELECT CASE WHEN SUM(r.TotalAmt - r.TotalPaid) > 0 THEN SUM(r.TotalAmt - r.TotalPaid) ELSE 0 END
                              FROM dbo.RegistrationTransactionSummary r
                              WHERE r.OrganizationId = om.OrganizationId AND r.PeopleId = om.PeopleId
                                AND r.IsLatestTransaction = 1), 0)
                 ELSE CASE WHEN (ISNULL((SELECT TOP 1 ts.IndDue FROM dbo.TransactionSummary ts
                                         WHERE ts.PeopleId = om.PeopleId AND ts.OrganizationId = om.OrganizationId
                                           AND ts.IsLatestTransaction = 1 ORDER BY ts.RegId DESC), 0) - ISNULL(sup.s, 0)) > 0
                           THEN (ISNULL((SELECT TOP 1 ts.IndDue FROM dbo.TransactionSummary ts
                                         WHERE ts.PeopleId = om.PeopleId AND ts.OrganizationId = om.OrganizationId
                                           AND ts.IsLatestTransaction = 1 ORDER BY ts.RegId DESC), 0) - ISNULL(sup.s, 0))
                           ELSE 0 END
            END AS Modal,
            ISNULL(sup.s, 0) AS Supporters
        FROM dbo.OrganizationMembers om
        JOIN dbo.Organizations o ON o.OrganizationId = om.OrganizationId
        JOIN dbo.People p ON p.PeopleId = om.PeopleId
        LEFT JOIN (SELECT GoerId, OrgId, SUM(Amount) s FROM dbo.GoerSenderAmounts
                   WHERE SupporterId <> GoerId GROUP BY GoerId, OrgId) sup
              ON sup.GoerId = om.PeopleId AND sup.OrgId = om.OrganizationId
        WHERE om.InactiveDate IS NULL AND om.TranId IS NOT NULL
    """
    # Level 2: clamp each surface to MAX(0, ...) -- the transactions screen, PayLink,
    # SB condition, and modal all display/treat a negative (overpaid) balance as $0.
    clamped = """
        SELECT OrgId, OrganizationName, RegType, Trip, PeopleId, TranId, Name2,
            CASE WHEN LedgerRaw  > 0 THEN LedgerRaw  ELSE 0 END AS Ledger,
            CASE WHEN PayLinkRaw > 0 THEN PayLinkRaw ELSE 0 END AS PayLink,
            CASE WHEN SBRaw      > 0 THEN SBRaw      ELSE 0 END AS SB,
            Modal, Supporters
        FROM ( %s ) r
    """ % (inner,)
    # Level 3: keep only rows where the (clamped) surfaces disagree.
    return """
    SELECT * FROM ( %s ) x
    WHERE (ABS(x.Ledger - x.PayLink) > 0.01 OR ABS(x.Ledger - x.SB) > 0.01
        OR ABS(x.Ledger - x.Modal) > 0.01 OR ABS(x.PayLink - x.SB) > 0.01) %s
    ORDER BY x.Trip DESC, x.OrganizationName, x.Name2
    """ % (clamped, org_filter)


def paylink_url(people_id, org_id):
    # model.GetPayLink returns the PayLink2 URL (no lt), which for new forms REDIRECTS to the
    # OnlineReg page and shows a recomputed $0. Appending &lt=1 makes it the StandardPayLink (the
    # member-dialog "PayLink" button), which skips the redirect and shows the AmountDueTrans balance,
    # matching the PayLink column.
    try:
        u = model.GetPayLink(people_id, org_id)
        if not u:
            return None
        return u + ('&lt=1' if '?' in u else '?lt=1')
    except:
        return None


def fetch(org_filter=""):
    out = []
    for r in q.QuerySql(rows_sql(org_filter)):
        out.append({
            'OrgId': r.OrgId,
            'OrgName': r.OrganizationName or ('Org ' + str(r.OrgId)),
            'IsNewForm': (r.RegType == 26),
            'IsTrip': (int(r.Trip or 0) == 1),
            'PeopleId': r.PeopleId,
            'TranId': r.TranId,
            'Name2': r.Name2 or ('Person ' + str(r.PeopleId)),
            'Ledger': float(r.Ledger or 0),
            'PayLink': float(r.PayLink or 0),
            'SB': float(r.SB or 0),
            'Modal': float(r.Modal or 0),
            'Supporters': float(r.Supporters or 0),
        })
    return out


# Distinct issue: negative TotalFee / TotPaid on the latest summary row. Set-based
# (joins the latest TransactionSummary row directly), so it stays fast.
def negatives_sql(org_filter=""):
    return """
    SELECT om.OrganizationId AS OrgId, o.OrganizationName, om.PeopleId, om.TranId AS TranId, p.Name2,
        ts.TotalFee, ts.TotPaid, ts.TotDue
    FROM dbo.OrganizationMembers om
    JOIN dbo.Organizations o ON o.OrganizationId = om.OrganizationId
    JOIN dbo.People p ON p.PeopleId = om.PeopleId
    JOIN dbo.TransactionSummary ts ON ts.PeopleId = om.PeopleId
        AND ts.OrganizationId = om.OrganizationId AND ts.IsLatestTransaction = 1
    WHERE om.InactiveDate IS NULL AND (ts.TotalFee < 0 OR ts.TotPaid < 0) %s
    ORDER BY o.OrganizationName, p.Name2
    """ % (org_filter,)


def fetch_negatives(org_filter=""):
    out, seen = [], set()
    for r in q.QuerySql(negatives_sql(org_filter)):
        key = (r.OrgId, r.PeopleId)
        if key in seen:
            continue
        seen.add(key)
        out.append({
            'OrgId': r.OrgId,
            'OrgName': r.OrganizationName or ('Org ' + str(r.OrgId)),
            'PeopleId': r.PeopleId,
            'TranId': r.TranId,
            'Name2': r.Name2 or ('Person ' + str(r.PeopleId)),
            'Fee': float(r.TotalFee or 0),
            'Paid': float(r.TotPaid or 0),
            'Due': float(r.TotDue or 0),
        })
    return out


def money(v):
    try:
        return "${:,.2f}".format(float(v or 0))
    except:
        return "$0.00"


def typebadges(x):
    b = '<span class="{0}">{1}</span>'.format('type nf' if x['IsNewForm'] else 'type',
                                               'New form' if x['IsNewForm'] else 'Classic')
    if x['IsTrip']:
        b += ' <span class="type trip">Trip</span>'
    return b


def d(a, b):
    return abs(a - b) > 0.01


# Category key, section title, and a one-line explanation for TouchPoint.
CATS = [
    ('trip', 'Mission-trip supporter accounting',
     'Supporter gifts (GoerSenderAmounts) and fee adjustments post differently to each surface, so the '
     'transaction ledger, PayLink, "Has Balance" condition, and involvement Due line do not reconcile. '
     'The ledger often goes negative (over-applied) while PayLink/SB keep a stale positive.'),
    ('rts_phantom', 'New-form summary shows a phantom balance (dialog blank)',
     'RegistrationTransactionSummary (Modal Due) reports a balance the transaction ledger, PayLink, and '
     '"Has Balance" condition all show as $0 (its TotalPaid does not credit a payment made by a family payer). '
     'The involvement dialog shows NO Fee/Paid/Due totals because the new-form total API errors on the mismatch, '
     'even though the summary still holds this amount.'),
    ('rts_missing', 'New-form Due shows $0 but a balance exists',
     'The transaction ledger, PayLink, and "Has Balance" condition show a balance, but the new-form involvement '
     'Due line (RegistrationTransactionSummary) reports $0, so the real balance is hidden on the involvement dialog.'),
    ('multiperson', 'Multi-person registration (transaction total vs individual share)',
     'The registration transaction covers more than one registrant. The transactions screen and the PayLink '
     'show the combined transaction balance (TotDue), while the "Has Balance" condition and the involvement '
     'Due line show only this person\'s share (IndDue).'),
    ('paylink', 'PayLink shows a stale balance',
     'The member-dialog "PayLink" button (StandardPayLink, lt=1) runs PaymentForm.AmountDueTrans, which returns '
     'the TotDue of the first TransactionSummary row with TotDue>0 for the registration, without filtering to the '
     'settled latest transaction. So it can show an old balance for someone who has paid (or a co-registrant\'s '
     'balance in a family registration) even though the ledger, "Has Balance" condition, and Due line all show $0.'),
    ('sb_stale', 'Has Balance condition reads a stale row',
     'The "Has Balance in Involvement" Search Builder condition reads TransactionSummary.IndDue of the '
     'highest-RegId row, not the latest transaction. Here that is a stale row, so the condition disagrees with '
     'the ledger, PayLink, and involvement Due line. This is the Search Builder over-reporting the church reported.'),
    ('other', 'Other surface disagreements',
     'Two or more of the four balance calculations return different values for the same registration.'),
]


def classify(x):
    L, P, S, M = x['Ledger'], x['PayLink'], x['SB'], x['Modal']
    if x['IsTrip']:
        return 'trip'
    # new-form RegistrationTransactionSummary (Modal) disagreeing with the TransactionSummary
    # surfaces (SB / ledger) is the RTS-out-of-sync issue, regardless of the other columns.
    if x['IsNewForm'] and d(M, S):
        return 'rts_phantom' if M > S + 0.01 else 'rts_missing'
    # transaction-total surfaces (ledger, paylink) agree, individual surfaces (SB, modal) agree, and they differ
    if not d(L, P) and not d(S, M) and d(L, S):
        return 'multiperson'
    if (d(P, S) or d(P, L)) and not d(S, L):
        return 'paylink'
    # SB (Has Balance) is the lone outlier: ledger, PayLink, and Modal agree but SB differs
    # (it reads the highest-RegId row, not the latest).
    if not d(L, P) and not d(L, M) and d(S, L):
        return 'sb_stale'
    return 'other'


CSS = """
<style>
  .pba { font-family:'Segoe UI',-apple-system,Arial,sans-serif; color:#1e293b; max-width:1180px; line-height:1.5; }
  .pba h2 { color:#1f4e79; margin:0 0 2px; font-size:24px; }
  .pba .lede { color:#64748b; font-size:13px; margin:0 0 16px; max-width:900px; }
  .pba code { background:#eef2f7; padding:1px 5px; border-radius:3px; font-size:11px; }

  .pba .cards { display:flex; gap:12px; flex-wrap:wrap; margin:0 0 18px; }
  .pba .card { flex:1; min-width:180px; border:1px solid #e2e8f0; border-radius:10px; padding:12px 14px; background:#fff; border-top:3px solid #1f4e79; }
  .pba .card .lbl { font-size:11px; text-transform:uppercase; letter-spacing:.6px; color:#64748b; font-weight:600; }
  .pba .card .val { font-size:24px; font-weight:800; margin-top:3px; color:#1f4e79; font-variant-numeric:tabular-nums; }
  .pba .card .sub { font-size:11px; color:#94a3b8; margin-top:2px; }

  .pba .surfaces { background:#f8fafc; border:1px solid #e2e8f0; border-left:4px solid #1f4e79;
                   border-radius:8px; padding:12px 16px; margin:0 0 18px; font-size:12px; color:#334155; }
  .pba .surfaces div { margin:3px 0; }
  .pba .surfaces b { color:#1e293b; }

  .pba .sec { margin:26px 0 4px; display:flex; align-items:baseline; gap:10px; }
  .pba .sec h3 { margin:0; font-size:16px; color:#1f4e79; }
  .pba .sec .count { color:#94a3b8; font-size:13px; }
  .pba .secx { font-size:12px; color:#64748b; margin:0 0 6px; max-width:900px; }

  .pba table { border-collapse:collapse; width:100%; font-size:13px; margin-top:2px; }
  .pba th { background:#1f4e79; color:#fff; text-align:left; padding:7px 9px; font-weight:600; font-size:12px; }
  .pba td { padding:6px 9px; border-bottom:1px solid #eef2f7; }
  .pba tr:hover td { background:#f5f9ff; }
  .pba .num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
  .pba .hi { background:#fff4f4; color:#b91c1c; font-weight:700; }
  .pba a { color:#1f4e79; text-decoration:none; }
  .pba a:hover { text-decoration:underline; }
  .pba .back { display:inline-block; margin-bottom:6px; font-size:13px; }
  .pba .muted { color:#94a3b8; }
  .pba .type { display:inline-block; font-size:11px; padding:1px 7px; border-radius:10px; background:#eef2f7; color:#475569; }
  .pba .type.nf { background:#fef3c7; color:#92400e; }
  .pba .type.trip { background:#e0e7ff; color:#3730a3; }
  .pba .btn { display:inline-block; background:#1f4e79; color:#fff; padding:7px 14px; border-radius:6px; font-size:13px; border:0; cursor:pointer; margin-bottom:14px; }
  .pba .btn:hover { background:#163a5a; }
  .pba .empty { background:#dcfce7; border:1px solid #bbf7d0; border-radius:8px; padding:14px 16px; color:#15803d; }
</style>
"""


def surfaces_block():
    return (
        '<div class="surfaces">'
        '<div>The four numbers below are TouchPoint\'s own calculations. A row appears only when they disagree.</div>'
        '<div><b>Ledger</b> = <code>TransactionSummary.TotDue</code> of the latest transaction (the transactions screen total; whole transaction, nets coupons/adjustments).</div>'
        '<div><b>PayLink</b> = the member-dialog "PayLink" button (StandardPayLink, <code>lt=1</code>) = '
        '<code>PaymentForm.AmountDueTrans</code>: the <code>TotDue</code> of the first TransactionSummary row with '
        'TotDue&gt;0 for the registration. (The email / regid paylink instead redirects and recomputes to $0 &mdash; '
        'so the two paylinks for the same person can disagree, which is itself a defect.)</div>'
        '<div><b>Has Balance (SB)</b> = <code>TransactionSummary.IndDue</code> of the highest-RegId row (the Search Builder condition).</div>'
        '<div><b>Modal Due</b> = the involvement pop-up "Due". Classic = individual due minus supporter gifts. '
        'New form = <code>RegistrationTransactionSummary</code> (the dialog\'s data source; the dialog itself can render '
        'BLANK when the new-form total API errors, even though the summary holds this value).</div>'
        '</div>'
    )


def print_button():
    return '<button class="btn" onclick="pbaPrint()">Print / Save this view</button>'


def cell(val, differs, href=None):
    txt = money(val)
    if href:
        txt = '<a href="{0}" target="_blank">{1}</a>'.format(href, txt)
    return '<td class="num{0}">{1}</td>'.format(' hi' if differs else '', txt)


def mismatch_table(rows):
    html = ['<table><thead><tr><th>Person</th><th>Involvement</th><th>Type</th>'
            '<th class="num">Ledger</th><th class="num">PayLink</th>'
            '<th class="num">Has Balance</th><th class="num">Modal Due</th>'
            '<th class="num">Supporter gifts</th></tr></thead><tbody>']
    for x in rows:
        L, P, S, M = x['Ledger'], x['PayLink'], x['SB'], x['Modal']
        # highlight cells that differ from the majority value
        vals = [L, P, S, M]
        base = max(set([round(v, 2) for v in vals]), key=[round(v, 2) for v in vals].count)
        # deep links to each surface for fast verification
        if not x['TranId']:
            ledger_href = None
        elif x['IsTrip']:
            ledger_href = '/Transactions/{0}?goerid={1}'.format(x['TranId'], x['PeopleId'])
        else:
            ledger_href = '/Transactions/{0}'.format(x['TranId'])
        pay_href = paylink_url(x['PeopleId'], x['OrgId'])
        modal_href = '/Person2/{0}#tab-current'.format(x['PeopleId'])
        sup = money(x['Supporters']) if x['Supporters'] > 0.01 else '<span class="muted">-</span>'
        if x['Supporters'] > 0.01 and x['TranId']:
            sup = '<a href="/Transactions/{0}?goerid={1}" target="_blank">{2}</a>'.format(x['TranId'], x['PeopleId'], money(x['Supporters']))
        html.append('<tr><td><a href="/Person2/{0}#enrollment" target="_blank">{1}</a></td>'
                    '<td><a href="/Organization/{2}" target="_blank">{3}</a></td>'
                    '<td>{4}</td>{5}{6}{7}{8}<td class="num muted">{9}</td></tr>'
                    .format(x['PeopleId'], x['Name2'], x['OrgId'], x['OrgName'], typebadges(x),
                            cell(L, d(L, base), ledger_href), cell(P, d(P, base), pay_href),
                            cell(S, d(S, base)), cell(M, d(M, base), modal_href), sup))
    html.append('</tbody></table>')
    return ''.join(html)


def render_sections(rows):
    buckets = {}
    for x in rows:
        buckets.setdefault(classify(x), []).append(x)
    html = []
    for key, title, expl in CATS:
        grp = buckets.get(key, [])
        if not grp:
            continue
        html.append('<div class="sec"><h3>{0}</h3><span class="count">{1} registrations</span></div>'.format(title, len(grp)))
        html.append('<p class="secx">{0}</p>'.format(expl))
        html.append(mismatch_table(grp))
    return ''.join(html)


def neg_table(rows):
    html = ['<table><thead><tr><th>Person</th><th>Involvement</th>'
            '<th class="num">Fee</th><th class="num">Paid</th><th class="num">Due</th></tr></thead><tbody>']
    for x in rows:
        due = money(x['Due'])
        if x['TranId']:
            due = '<a href="/Transactions/{0}" target="_blank">{1}</a>'.format(x['TranId'], due)
        html.append('<tr><td><a href="/Person2/{0}#tab-current" target="_blank">{1}</a></td>'
                    '<td><a href="/Organization/{2}" target="_blank">{3}</a></td>'
                    '<td class="num hi">{4}</td><td class="num hi">{5}</td><td class="num">{6}</td></tr>'
                    .format(x['PeopleId'], x['Name2'], x['OrgId'], x['OrgName'],
                            money(x['Fee']), money(x['Paid']), due))
    html.append('</tbody></table>')
    return ''.join(html)


def render_negatives(negs):
    if not negs:
        return ''
    html = ['<div class="sec"><div class="bar" style="background:#b45309;"></div>'
            '<h3>Negative fee / paid (recurring-billing artifact)</h3>'
            '<span class="count">{0} registrations</span></div>'.format(len(negs))]
    html.append('<p class="secx"><b>Distinct issue, not a surface conflict.</b> Recurring monthly billing '
                'accumulates a negative <code>TotalFee</code> / <code>TotPaid</code> on the latest summary row. Due '
                'still nets to $0, so it reads correctly day-to-day, but a manual fee adjustment breaks the math: '
                'adding $200 to a -$200 fee leaves Paid at -$200, so Due flips to +$200 instead of $0. Almost all are '
                'the recurring Payers plans.</p>')
    html.append(neg_table(negs))
    return ''.join(html)


def render_full(rows, negs):
    trip = [x for x in rows if x['IsTrip']]
    nontrip = [x for x in rows if not x['IsTrip']]
    html = []
    html.append('<div class="cards">')
    html.append('<div class="card"><div class="lbl">Surface conflicts</div>'
                '<div class="val">{0}</div><div class="sub">the four balances disagree</div></div>'.format(len(rows)))
    html.append('<div class="card"><div class="lbl">Mission-trip conflicts</div>'
                '<div class="val">{0}</div><div class="sub">supporter-gift accounting</div></div>'.format(len(trip)))
    html.append('<div class="card"><div class="lbl">Non-trip conflicts</div>'
                '<div class="val">{0}</div><div class="sub">new-form sync / multi-person</div></div>'.format(len(nontrip)))
    html.append('<div class="card"><div class="lbl">Negative fee / paid</div>'
                '<div class="val">{0}</div><div class="sub">recurring-billing data anomaly</div></div>'.format(len(negs)))
    html.append('</div>')
    html.append(surfaces_block())
    html.append(print_button())
    if not rows:
        html.append('<div class="empty">No surface conflicts found. Every registration\'s four balances agree.</div>')
    else:
        html.append(render_sections(rows))
    html.append(render_negatives(negs))
    return ''.join(html)


def render_drill(org_id):
    org_i = int(org_id)
    rows = fetch("AND x.OrgId = " + str(org_i))
    negs = fetch_negatives("AND om.OrganizationId = " + str(org_i))
    oname = rows[0]['OrgName'] if rows else (negs[0]['OrgName'] if negs else ('Org ' + str(org_i)))
    html = ['<a class="back" href="?">&larr; Back to full audit</a>']
    html.append('<h2>{0} <span class="muted" style="font-weight:400;font-size:15px;">(org {1})</span></h2>'.format(oname, org_i))
    html.append(surfaces_block())
    html.append(print_button())
    if rows:
        html.append(render_sections(rows))
    else:
        html.append('<div class="empty">No surface conflicts in this involvement.</div>')
    html.append(render_negatives(negs))
    return ''.join(html)


PRINT_JS = """
<script>
function pbaPrint(){
  var node = document.querySelector('.pba');
  if(!node){ window.print(); return; }
  var clone = node.cloneNode(true);
  var btns = clone.querySelectorAll('.btn'); for(var i=0;i<btns.length;i++){ btns[i].parentNode.removeChild(btns[i]); }
  var css='';
  css+='*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}';
  css+="body{margin:0;padding:24px;font-family:'Segoe UI',Arial,sans-serif;color:#1e293b}";
  css+='h2{color:#1f4e79;margin:0 0 4px}h3{margin:0;color:#1f4e79}';
  css+='.surfaces{background:#f8fafc;border:1px solid #e2e8f0;border-left:4px solid #1f4e79;border-radius:8px;padding:10px 14px;margin:0 0 16px;font-size:11px;color:#334155}';
  css+='.cards{display:flex;gap:12px;margin-bottom:16px}';
  css+='.card{flex:1;border:1px solid #e2e8f0;border-radius:10px;padding:12px 14px;border-top:3px solid #1f4e79}';
  css+='.card .lbl{font-size:10px;text-transform:uppercase;letter-spacing:.6px;color:#64748b;font-weight:600}';
  css+='.card .val{font-size:22px;font-weight:800;margin-top:3px;color:#1f4e79}.card .sub{font-size:11px;color:#94a3b8}';
  css+='.sec{margin:18px 0 4px}.sec h3{color:#1f4e79}.sec .count{color:#94a3b8;font-size:12px;margin-left:8px}.secx{font-size:11px;color:#64748b;margin:0 0 6px}';
  css+='table{border-collapse:collapse;width:100%;font-size:11px}';
  css+='th{background:#1f4e79;color:#fff;text-align:left;padding:6px 8px;font-size:10px}';
  css+='td{padding:5px 8px;border-bottom:1px solid #eef2f7}';
  css+='.num{text-align:right}.hi{background:#fff4f4;color:#b91c1c;font-weight:700}.muted{color:#94a3b8}';
  css+='.type{font-size:10px;padding:1px 7px;border-radius:10px;background:#eef2f7;color:#475569}.type.nf{background:#fef3c7;color:#92400e}.type.trip{background:#e0e7ff;color:#3730a3}';
  css+='.back{display:none}a{color:#1e293b;text-decoration:none}';
  var pw=window.open('','_blank');
  if(!pw){ alert('Popup blocked - please allow popups to print'); return; }
  pw.document.write('<!DOCTYPE html><html><head><title>Balance Consistency Audit</title><style>'+css+'</style></head><body>');
  pw.document.write(clone.innerHTML);
  pw.document.write('</body></html>');
  pw.document.close(); pw.focus();
  setTimeout(function(){ pw.print(); }, 350);
}
</script>
"""


# ------------------------------------------------------------------ #
org_param = None
try:
    if hasattr(model.Data, 'org') and str(model.Data.org).strip():
        org_param = str(model.Data.org).strip()
except:
    org_param = None

print CSS
print PRINT_JS
print '<div class="pba">'
print '<h2>Registration Balance Consistency Audit</h2>'
print ('<p class="lede">Every registration where TouchPoint\'s four balance calculations '
       '(ledger, PayLink, "Has Balance" condition, and involvement Due line) disagree with each other. '
       'Evidence for TouchPoint support, not a list of who owes. Read-only.</p>')

try:
    if org_param:
        print render_drill(org_param)
    else:
        print render_full(fetch(), fetch_negatives())
except Exception as e:
    print '<div style="color:#b91c1c;font-family:monospace;white-space:pre-wrap;">Error: {0}\n{1}</div>'.format(str(e), traceback.format_exc())

print '</div>'
