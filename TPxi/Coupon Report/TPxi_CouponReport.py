#roles=Edit,Finance
#----------------------------------------------------------------------
# TPxi_CouponReport.py
#
# Post-mortem report for coupon / discount code usage across involvements.
#
#
# Output:
#   - Summary tiles (Used / $ Discounted / Unused / Cancelled)
#   - "By Involvement" roll-up
#   - "By Name on Code" roll-up (the scholarship-fund pivot)
#   - Detail grid (every coupon row)
#   - CSV export of the detail grid
#
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

import json
import datetime
import traceback

model.Header = 'Coupon Report'

APP_VERSION = '1.0.0'
DC_SCRIPT_ID = 'TPxi_CouponReport'


# ===== Helpers =====

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
    if s is None:
        return ''
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
        return dict((sanitize_for_json(k), sanitize_for_json(v)) for k, v in obj.items())
    if isinstance(obj, (list, tuple)):
        return [sanitize_for_json(item) for item in obj]
    return safe_str(obj)


def jprint(obj):
    import sys
    sys.stdout.write(json.dumps(sanitize_for_json(obj)))


def safe_int(v, default=0):
    try:
        return int(v)
    except:
        return default


def safe_money(v):
    try:
        if v is None:
            return 0.0
        return float(v)
    except:
        return 0.0


def get_data(name, default=''):
    try:
        if hasattr(model.Data, name):
            v = getattr(model.Data, name)
            return v if v is not None else default
    except:
        pass
    return default


def safe_date(v):
    """Format a SQL date/datetime as YYYY-MM-DD."""
    if v is None:
        return ''
    try:
        return v.strftime('%Y-%m-%d')
    except:
        try:
            s = safe_str(v)
            return s[:10] if s else ''
        except:
            return ''


# ===== Scope resolution =====

def resolve_scope_orgs(prog_div_groups, specific_org_ids):
    """Active orgs in scope (union of program/division rows + specific orgs)."""
    org_ids = set()
    for oid in specific_org_ids:
        oid = safe_int(oid, 0)
        if oid > 0:
            org_ids.add(oid)
    for g in prog_div_groups:
        prog_id = safe_int(g.get('programId'), 0)
        div_id = safe_int(g.get('divisionId'), 0)
        if prog_id <= 0 and div_id <= 0:
            continue
        wh = ["o.OrganizationStatusId = 30"]
        if prog_id > 0:
            wh.append("os.ProgId = " + str(prog_id))
        if div_id > 0:
            wh.append("os.DivId = " + str(div_id))
        sql = ("SELECT DISTINCT o.OrganizationId FROM Organizations o "
               "JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId "
               "WHERE " + " AND ".join(wh))
        try:
            for r in q.QuerySql(sql):
                org_ids.add(r.OrganizationId)
        except:
            pass
    return list(org_ids)


# ===== AJAX handlers =====

def handle_get_filters():
    progs = []
    for r in q.QuerySql("SELECT Id, Name FROM Program WHERE Name IS NOT NULL ORDER BY Name"):
        progs.append({'id': r.Id, 'name': safe_str(r.Name)})
    divs = []
    for r in q.QuerySql("SELECT Id, Name, ProgId FROM Division WHERE Name IS NOT NULL ORDER BY Name"):
        divs.append({'id': r.Id, 'name': safe_str(r.Name), 'programId': r.ProgId})
    jprint({'success': True, 'programs': progs, 'divisions': divs})


def handle_search_involvements():
    term = get_data('cr_term', '').strip()
    if not term:
        jprint({'success': True, 'results': []})
        return
    safe_term = term.replace("'", "''")
    sql = ("SELECT TOP 30 o.OrganizationId, o.OrganizationName, "
           "ISNULL(MAX(os.Division), '') AS Division, "
           "ISNULL(MAX(os.Program), '') AS Program "
           "FROM Organizations o "
           "LEFT JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId "
           "WHERE o.OrganizationStatusId = 30 "
           "AND o.OrganizationName LIKE '%" + safe_term + "%' "
           "GROUP BY o.OrganizationId, o.OrganizationName "
           "ORDER BY o.OrganizationName")
    results = []
    for r in q.QuerySql(sql):
        results.append({
            'orgId': r.OrganizationId,
            'orgName': safe_str(r.OrganizationName),
            'division': safe_str(r.Division),
            'program': safe_str(r.Program),
        })
    jprint({'success': True, 'results': results})


def handle_bulk_lookup_involvements():
    import re
    raw = get_data('cr_ids_raw', '')
    ids = []
    seen = set()
    for tok in re.split(r'[^0-9]+', raw):
        if not tok:
            continue
        try:
            n = int(tok)
        except:
            continue
        if n > 0 and n not in seen:
            seen.add(n)
            ids.append(n)
    if not ids:
        jprint({'success': True, 'found': [], 'missing': [], 'inactive': []})
        return
    id_csv = ','.join(str(x) for x in ids)
    sql = ("SELECT o.OrganizationId, o.OrganizationName, o.OrganizationStatusId, "
           "ISNULL(MAX(CAST(os.Division AS NVARCHAR(200))), '') AS Division, "
           "ISNULL(MAX(CAST(os.Program AS NVARCHAR(200))), '') AS Program "
           "FROM Organizations o "
           "LEFT JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId "
           "WHERE o.OrganizationId IN (" + id_csv + ") "
           "GROUP BY o.OrganizationId, o.OrganizationName, o.OrganizationStatusId")
    found = []
    inactive = []
    seen_ids = set()
    for r in q.QuerySql(sql):
        oid = r.OrganizationId
        seen_ids.add(int(oid))
        entry = {
            'orgId': oid,
            'orgName': safe_str(r.OrganizationName),
            'division': safe_str(r.Division),
            'program': safe_str(r.Program),
        }
        if r.OrganizationStatusId == 30:
            found.append(entry)
        else:
            inactive.append(entry)
    missing = [i for i in ids if i not in seen_ids]
    jprint({'success': True, 'found': found, 'missing': missing, 'inactive': inactive})


def handle_run_report():
    """Pull coupons in scope, return summary + roll-ups + detail rows.

    Filters:
      - cr_prog_div_groups : JSON list of {programId, divisionId}
      - cr_specific_org_ids: JSON list of OrgIds
      - cr_date_field      : 'used' | 'created'   (default 'used')
      - cr_date_from       : YYYY-MM-DD inclusive
      - cr_date_to         : YYYY-MM-DD inclusive
      - cr_status          : 'used' | 'unused' | 'cancelled' | 'all' (default 'used')
      - cr_code_prefix     : LIKE pattern prefix on Coupons.Id
      - cr_name_search     : LIKE pattern on Coupons.Name
    """
    try:
        prog_div_groups = json.loads(get_data('cr_prog_div_groups', '[]'))
    except:
        prog_div_groups = []
    try:
        specific_org_ids = json.loads(get_data('cr_specific_org_ids', '[]'))
    except:
        specific_org_ids = []
    date_field = (get_data('cr_date_field', 'used') or 'used').strip().lower()
    if date_field not in ('used', 'created'):
        date_field = 'used'
    date_from = (get_data('cr_date_from', '') or '').strip()
    date_to = (get_data('cr_date_to', '') or '').strip()
    status = (get_data('cr_status', 'used') or 'used').strip().lower()
    if status not in ('used', 'unused', 'cancelled', 'all'):
        status = 'used'
    code_prefix = (get_data('cr_code_prefix', '') or '').strip()
    name_search = (get_data('cr_name_search', '') or '').strip()

    scope_orgs = resolve_scope_orgs(prog_div_groups, specific_org_ids)
    if not scope_orgs:
        jprint({
            'success': True,
            'message': 'Empty scope. Add a program, division, or specific involvement.',
            'summary': {'total': 0, 'totalDiscounted': 0.0, 'used': 0, 'unused': 0, 'cancelled': 0},
            'rows': [], 'byOrg': [], 'byName': []
        })
        return

    where = []
    where.append("c.OrgId IN (" + ",".join(str(x) for x in scope_orgs) + ")")

    if status == 'used':
        where.append("c.Used IS NOT NULL")
        where.append("c.Canceled IS NULL")
    elif status == 'unused':
        where.append("c.Used IS NULL")
        where.append("c.Canceled IS NULL")
    elif status == 'cancelled':
        where.append("c.Canceled IS NOT NULL")
    # 'all' adds no status filter

    # Date filter on the chosen field. Treat date_to inclusively by adding 23:59:59.
    date_col = 'c.Used' if date_field == 'used' else 'c.Created'
    if date_from:
        df = date_from.replace("'", "")
        where.append("{0} >= '{1}'".format(date_col, df))
    if date_to:
        dt = date_to.replace("'", "")
        where.append("{0} <= '{1} 23:59:59'".format(date_col, dt))

    if code_prefix:
        cp = code_prefix.replace("'", "''")
        where.append("c.Id LIKE '" + cp + "%'")
    if name_search:
        ns = name_search.replace("'", "''")
        where.append("c.Name LIKE '%" + ns + "%'")

    where_sql = " AND ".join(where)

    detail_sql = (
        "SELECT TOP 5000 "
        "    c.Id AS Code, "
        "    ISNULL(c.Name, '') AS NameOnCode, "
        "    c.Amount, "
        "    c.RegAmount, "
        "    c.Used, "
        "    c.Created, "
        "    c.Canceled, "
        "    c.OrgId, "
        "    ISNULL(o.OrganizationName, '') AS OrgName, "
        "    c.PeopleId, "
        "    ISNULL(p.Name2, '') AS PersonName, "
        "    c.UserId, "
        "    ISNULL(up.Name2, '') AS CreatedByName, "
        "    ISNULL(c.Type, '') AS Type, "
        "    c.Percentage, "
        "    ISNULL(c.MultiUse, 0) AS MultiUse, "
        "    ISNULL(c.Generated, 0) AS Generated "
        "FROM Coupons c "
        "LEFT JOIN Organizations o ON o.OrganizationId = c.OrgId "
        "LEFT JOIN People p ON p.PeopleId = c.PeopleId "
        "LEFT JOIN People up ON up.PeopleId = c.UserId "
        "WHERE " + where_sql + " "
        "ORDER BY ISNULL(c.Used, c.Created) DESC, c.Id"
    )

    rows = []
    total_disc = 0.0
    cnt_used = 0
    cnt_unused = 0
    cnt_cancel = 0

    by_org = {}    # OrgId -> {orgName, count, total}
    by_name = {}   # NameOnCode -> {count, total}

    for r in q.QuerySql(detail_sql):
        used_dt = r.Used
        canceled_dt = r.Canceled
        if canceled_dt is not None:
            row_status = 'Cancelled'
            cnt_cancel += 1
        elif used_dt is not None:
            row_status = 'Used'
            cnt_used += 1
        else:
            row_status = 'Unused'
            cnt_unused += 1

        amt = safe_money(r.Amount)
        # Only count toward "discounted" if used and not cancelled.
        if row_status == 'Used':
            total_disc += amt

        org_id = safe_int(r.OrgId, 0)
        org_name = safe_str(r.OrgName) or ('Org #' + str(org_id) if org_id else '')
        name_on_code = safe_str(r.NameOnCode).strip()
        name_key = name_on_code if name_on_code else '(no name)'

        rows.append({
            'code': safe_str(r.Code),
            'nameOnCode': name_on_code,
            'amount': amt,
            'regAmount': safe_money(r.RegAmount),
            'used': safe_date(used_dt),
            'created': safe_date(r.Created),
            'cancelled': safe_date(canceled_dt),
            'orgId': org_id,
            'orgName': org_name,
            'peopleId': safe_int(r.PeopleId, 0),
            'personName': safe_str(r.PersonName),
            'createdById': safe_int(r.UserId, 0),
            'createdByName': safe_str(r.CreatedByName),
            'type': safe_str(r.Type),
            'percentage': safe_money(r.Percentage),
            'multiUse': bool(int(r.MultiUse or 0)),
            'generated': bool(int(r.Generated or 0)),
            'status': row_status,
        })

        if org_id:
            o = by_org.setdefault(org_id, {'orgId': org_id, 'orgName': org_name,
                                            'used': 0, 'unused': 0, 'cancelled': 0,
                                            'totalDiscounted': 0.0})
            if row_status == 'Used':
                o['used'] += 1
                o['totalDiscounted'] += amt
            elif row_status == 'Unused':
                o['unused'] += 1
            else:
                o['cancelled'] += 1

        n = by_name.setdefault(name_key, {'nameOnCode': name_key,
                                            'used': 0, 'unused': 0, 'cancelled': 0,
                                            'totalDiscounted': 0.0})
        if row_status == 'Used':
            n['used'] += 1
            n['totalDiscounted'] += amt
        elif row_status == 'Unused':
            n['unused'] += 1
        else:
            n['cancelled'] += 1

    by_org_list = sorted(by_org.values(), key=lambda x: x['totalDiscounted'], reverse=True)
    by_name_list = sorted(by_name.values(), key=lambda x: x['totalDiscounted'], reverse=True)

    jprint({
        'success': True,
        'summary': {
            'total': len(rows),
            'totalDiscounted': round(total_disc, 2),
            'used': cnt_used,
            'unused': cnt_unused,
            'cancelled': cnt_cancel,
            'scopeOrgCount': len(scope_orgs),
        },
        'rows': rows,
        'byOrg': by_org_list,
        'byName': by_name_list,
        'truncated': len(rows) >= 5000,
    })


# ===== Dispatch =====

if model.HttpMethod == "post":
    action = get_data('cr_action', '')
    try:
        if action == 'get_filters':
            handle_get_filters()
        elif action == 'search_involvements':
            handle_search_involvements()
        elif action == 'bulk_lookup_involvements':
            handle_bulk_lookup_involvements()
        elif action == 'run_report':
            handle_run_report()
        else:
            jprint({'success': False, 'message': 'Unknown action: ' + safe_str(action)})
    except Exception as e:
        jprint({
            'success': False,
            'message': 'Server error: ' + safe_str(e),
            'trace': traceback.format_exc(),
        })
else:
    model.Form = """
<style>
.cr-root { font-family: 'Segoe UI', Arial, sans-serif; color: #222; max-width: 1500px; margin: 0 auto; padding: 12px; }
.cr-card { background: #fff; border: 1px solid #e1e4e8; border-radius: 8px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
.cr-h1 { font-size: 22px; font-weight: 700; margin: 0 0 4px 0; color: #1f4e79; }
.cr-h2 { font-size: 17px; font-weight: 700; margin: 0 0 8px 0; color: #1f4e79; }
.cr-sub { font-size: 13px; color: #555; margin-bottom: 8px; }
.cr-label { display: block; font-size: 12px; color: #555; margin: 10px 0 4px 0; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
.cr-muted { color: #888; font-size: 12px; }
.cr-row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
.cr-input, .cr-select { padding: 7px 10px; font-size: 14px; border: 1px solid #c8ccd0; border-radius: 4px; background: #fff; min-width: 180px; box-sizing: border-box; }
.cr-input:focus, .cr-select:focus { outline: none; border-color: #1f4e79; box-shadow: 0 0 0 2px rgba(31, 78, 121, 0.15); }
.cr-btn { display: inline-block; padding: 8px 14px; font-size: 14px; font-weight: 600; border: 1px solid #1f4e79; background: #1f4e79; color: #fff; border-radius: 4px; cursor: pointer; text-decoration: none; }
.cr-btn:hover { background: #2a5e8e; }
.cr-btn.cr-secondary { background: #fff; color: #1f4e79; }
.cr-btn.cr-secondary:hover { background: #f0f4f8; }
.cr-btn.cr-sm { padding: 4px 9px; font-size: 12px; }
.cr-btn.cr-danger { background: #c0392b; border-color: #c0392b; }
.cr-btn:disabled { opacity: 0.5; cursor: not-allowed; }
.cr-progdiv-rows { display: flex; flex-direction: column; gap: 6px; margin-bottom: 8px; }
.cr-progdiv-row { display: flex; gap: 6px; align-items: center; }
.cr-chip { display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px; background: #eef2f7; border: 1px solid #c8d4e3; border-radius: 999px; font-size: 13px; }
.cr-chip-x { cursor: pointer; color: #c0392b; font-weight: 700; padding: 0 2px; }
.cr-chip-row { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 4px; }
.cr-search-wrap { position: relative; }
.cr-search-results { max-height: 320px; overflow-y: auto; border: 1px solid #e1e4e8; border-radius: 4px; margin-top: 4px; background: #fff; box-shadow: 0 4px 12px rgba(0,0,0,0.1); position: absolute; z-index: 100; width: 100%; }
.cr-search-result { padding: 6px 10px; cursor: pointer; border-bottom: 1px solid #f0f0f0; }
.cr-search-result:hover { background: #f0f4f8; }
.cr-tile-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 14px; }
.cr-tile { flex: 1 1 160px; background: #fff; border: 1px solid #e1e4e8; border-left: 4px solid #1f4e79; border-radius: 6px; padding: 10px 14px; min-width: 150px; }
.cr-tile-num { font-size: 22px; font-weight: 700; color: #1f4e79; line-height: 1.1; }
.cr-tile-label { font-size: 11px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 2px; }
.cr-tile.cr-tile-money { border-left-color: #27ae60; }
.cr-tile.cr-tile-money .cr-tile-num { color: #27ae60; }
.cr-tile.cr-tile-unused { border-left-color: #f39c12; }
.cr-tile.cr-tile-unused .cr-tile-num { color: #b67900; }
.cr-tile.cr-tile-cancelled { border-left-color: #c0392b; }
.cr-tile.cr-tile-cancelled .cr-tile-num { color: #c0392b; }
.cr-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.cr-table th { text-align: left; padding: 8px 10px; background: #f0f4f8; font-weight: 700; color: #1f4e79; border-bottom: 2px solid #d8e0e8; white-space: nowrap; }
.cr-table td { padding: 7px 10px; border-bottom: 1px solid #eee; vertical-align: top; }
.cr-table tr:hover td { background: #fafbfc; }
.cr-table th.num, .cr-table td.num { text-align: right; }
.cr-status-used { background: #dff6dd; color: #107c10; padding: 1px 7px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.cr-status-unused { background: #fff4ce; color: #7a5c00; padding: 1px 7px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.cr-status-cancelled { background: #fde7e9; color: #c0392b; padding: 1px 7px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.cr-subtabs { display: flex; gap: 0; border-bottom: 2px solid #e1e4e8; margin-bottom: 10px; }
.cr-subtab { padding: 8px 16px; cursor: pointer; font-weight: 600; font-size: 14px; color: #888; border-bottom: 2px solid transparent; margin-bottom: -2px; }
.cr-subtab:hover { color: #1f4e79; }
.cr-subtab.active { color: #1f4e79; border-bottom-color: #1f4e79; }
.cr-empty { color: #888; padding: 30px; text-align: center; font-style: italic; }
.cr-back { color: #1f4e79; cursor: pointer; font-weight: 600; font-size: 14px; margin-bottom: 8px; display: inline-block; }
.cr-back:hover { text-decoration: underline; }
</style>

<div class="cr-root" id="crRoot">
  <div style="display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:8px; gap:12px;">
    <div>
      <div class="cr-h1">Coupon Report <span style="font-size:11px; color:#888; font-weight:400;">v""" + APP_VERSION + """</span></div>
      <div class="cr-sub">Track coupon / discount code usage across involvements. Preserves the "Name on Code" naming convention for fund reconciliation.</div>
    </div>
  </div>
  <div id="crMain"></div>
</div>

<script>
var APP_VERSION = """ + json.dumps(APP_VERSION) + """;

var state = {
  view: 'scope',
  programs: [],
  divisions: [],
  progDivGroups: [{programId: 0, divisionId: 0}],
  specificOrgs: [],
  dateField: 'used',
  dateFrom: '',
  dateTo: '',
  status: 'used',
  codePrefix: '',
  nameSearch: '',
  result: null,
  resultTab: 'byName',
};

// Default to last 12 months -> today.
(function setDefaultDates() {
  var now = new Date();
  var yr = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
  function fmt(d) {
    var m = String(d.getMonth() + 1); if (m.length === 1) m = '0' + m;
    var dd = String(d.getDate()); if (dd.length === 1) dd = '0' + dd;
    return d.getFullYear() + '-' + m + '-' + dd;
  }
  state.dateFrom = fmt(yr);
  state.dateTo = fmt(now);
})();

var root = document.getElementById('crMain');

function el(tag, attrs, content) {
  var n = document.createElement(tag);
  if (attrs) {
    for (var k in attrs) {
      if (k === 'class') n.className = attrs[k];
      else if (k === 'style') n.style.cssText = attrs[k];
      else if (k.indexOf('on') === 0) n[k] = attrs[k];
      else n.setAttribute(k, attrs[k]);
    }
  }
  if (content !== undefined && content !== null) {
    if (typeof content === 'string') n.textContent = content;
    else if (Array.isArray(content)) content.forEach(function(c) { if (c) n.appendChild(c); });
    else n.appendChild(content);
  }
  return n;
}

function fmtMoney(v) {
  v = Number(v || 0);
  return '$' + v.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function escapeHtml(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function ajax(actionName, data, cb) {
  var xhr = new XMLHttpRequest();
  xhr.open('POST', window.location.pathname, true);
  xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  var params = ['cr_action=' + encodeURIComponent(actionName)];
  for (var k in data) {
    if (data[k] === undefined || data[k] === null) continue;
    params.push(encodeURIComponent(k) + '=' + encodeURIComponent(data[k]));
  }
  xhr.onload = function() {
    if (xhr.status >= 200 && xhr.status < 300) {
      try { cb(null, JSON.parse(xhr.responseText)); }
      catch (e) { cb(e, null); }
    } else {
      cb(new Error('HTTP ' + xhr.status), null);
    }
  };
  xhr.onerror = function() { cb(new Error('Network error'), null); };
  xhr.send(params.join('&'));
}

function render() {
  root.innerHTML = '';
  if (state.view === 'scope') renderScope();
  else if (state.view === 'results') renderResults();
}

function renderScope() {
  var card = el('div', {class: 'cr-card'});
  card.appendChild(el('div', {class: 'cr-h2'}, 'Choose what to report on'));
  card.appendChild(el('div', {class: 'cr-sub'}, 'Scope is the UNION of every Program/Division row plus any specific involvements you add.'));

  // Programs & Divisions
  card.appendChild(el('div', {class: 'cr-label'}, 'Programs & Divisions'));
  var pdWrap = el('div', {class: 'cr-progdiv-rows', id: 'crPdRows'});
  card.appendChild(pdWrap);
  renderProgDivRows();
  card.appendChild(el('button', {class: 'cr-btn cr-secondary cr-sm', onclick: function() {
    state.progDivGroups.push({programId: 0, divisionId: 0});
    renderProgDivRows();
  }}, '+ Add Program/Division'));

  // Specific involvements
  card.appendChild(el('div', {class: 'cr-label'}, 'Specific Involvements (optional)'));
  var chipRow = el('div', {class: 'cr-chip-row', id: 'crSpecificChips'});
  card.appendChild(chipRow);
  renderSpecificChips();
  var searchWrap = el('div', {class: 'cr-search-wrap'});
  var searchInput = el('input', {
    class: 'cr-input', type: 'text', style: 'width:100%; max-width:520px;',
    placeholder: 'Search by involvement name or paste IDs...',
    oninput: function(e) {
      var v = e.target.value;
      if (/^[\\s0-9,;\\n\\r]+$/.test(v) && /\\d/.test(v) && /[,;\\s]/.test(v)) {
        bulkAddFromString(v, function() { e.target.value = ''; });
      } else {
        searchOrgs(v);
      }
    }
  });
  searchWrap.appendChild(searchInput);
  searchWrap.appendChild(el('div', {id: 'crOrgSearchResults'}));
  card.appendChild(searchWrap);

  // Filters
  card.appendChild(el('div', {class: 'cr-label'}, 'Date range'));
  var dateRow = el('div', {class: 'cr-row'});
  var dateFieldSel = el('select', {class: 'cr-select', style: 'min-width:160px;', onchange: function(e) {
    state.dateField = e.target.value;
  }});
  ['used', 'created'].forEach(function(v) {
    var label = v === 'used' ? 'Date Used' : 'Date Created';
    var opt = el('option', {value: v}, label);
    if (state.dateField === v) opt.selected = true;
    dateFieldSel.appendChild(opt);
  });
  dateRow.appendChild(dateFieldSel);
  var fromInput = el('input', {type: 'date', class: 'cr-input', style: 'min-width:160px;',
    value: state.dateFrom, onchange: function(e) { state.dateFrom = e.target.value; }});
  dateRow.appendChild(el('span', {class: 'cr-muted'}, 'from'));
  dateRow.appendChild(fromInput);
  var toInput = el('input', {type: 'date', class: 'cr-input', style: 'min-width:160px;',
    value: state.dateTo, onchange: function(e) { state.dateTo = e.target.value; }});
  dateRow.appendChild(el('span', {class: 'cr-muted'}, 'to'));
  dateRow.appendChild(toInput);
  card.appendChild(dateRow);

  card.appendChild(el('div', {class: 'cr-label'}, 'Status'));
  var statusSel = el('select', {class: 'cr-select', style: 'min-width:180px;', onchange: function(e) {
    state.status = e.target.value;
  }});
  [['used', 'Used (default)'], ['unused', 'Unused (issued, not used)'],
   ['cancelled', 'Cancelled'], ['all', 'All']].forEach(function(pair) {
    var opt = el('option', {value: pair[0]}, pair[1]);
    if (state.status === pair[0]) opt.selected = true;
    statusSel.appendChild(opt);
  });
  card.appendChild(statusSel);

  card.appendChild(el('div', {class: 'cr-label'}, 'Optional code/name filters'));
  var filterRow = el('div', {class: 'cr-row'});
  var prefixInput = el('input', {class: 'cr-input', type: 'text', placeholder: 'Code starts with...',
    value: state.codePrefix, oninput: function(e) { state.codePrefix = e.target.value; }});
  var nameInput = el('input', {class: 'cr-input', type: 'text', placeholder: 'Name on Code contains...',
    value: state.nameSearch, oninput: function(e) { state.nameSearch = e.target.value; }});
  filterRow.appendChild(prefixInput);
  filterRow.appendChild(nameInput);
  card.appendChild(filterRow);

  card.appendChild(el('div', {style: 'height:14px'}));
  var runBtn = el('button', {class: 'cr-btn', onclick: runReport}, 'Run Report');
  card.appendChild(runBtn);

  root.appendChild(card);
}

function renderProgDivRows() {
  var wrap = document.getElementById('crPdRows');
  if (!wrap) return;
  wrap.innerHTML = '';
  state.progDivGroups.forEach(function(g, idx) {
    var row = el('div', {class: 'cr-progdiv-row'});
    var pSel = el('select', {class: 'cr-select', onchange: function(e) {
      g.programId = parseInt(e.target.value);
      g.divisionId = 0;
      renderProgDivRows();
    }});
    pSel.appendChild(el('option', {value: 0}, '(any program)'));
    state.programs.forEach(function(p) {
      var opt = el('option', {value: p.id}, p.name);
      if (p.id === g.programId) opt.selected = true;
      pSel.appendChild(opt);
    });
    var dSel = el('select', {class: 'cr-select', onchange: function(e) {
      g.divisionId = parseInt(e.target.value);
    }});
    dSel.appendChild(el('option', {value: 0}, '(any division)'));
    state.divisions.filter(function(d) { return !g.programId || d.programId === g.programId; }).forEach(function(d) {
      var opt = el('option', {value: d.id}, d.name);
      if (d.id === g.divisionId) opt.selected = true;
      dSel.appendChild(opt);
    });
    var rm = el('button', {class: 'cr-btn cr-danger cr-sm', onclick: function() {
      state.progDivGroups.splice(idx, 1);
      if (!state.progDivGroups.length) state.progDivGroups.push({programId: 0, divisionId: 0});
      renderProgDivRows();
    }}, 'x');
    row.appendChild(pSel);
    row.appendChild(dSel);
    row.appendChild(rm);
    wrap.appendChild(row);
  });
}

function renderSpecificChips() {
  var wrap = document.getElementById('crSpecificChips');
  if (!wrap) return;
  wrap.innerHTML = '';
  state.specificOrgs.forEach(function(o, idx) {
    var chip = el('span', {class: 'cr-chip'});
    chip.appendChild(document.createTextNode(o.orgName + ' (#' + o.orgId + ')'));
    chip.appendChild(el('span', {class: 'cr-chip-x', onclick: function() {
      state.specificOrgs.splice(idx, 1);
      renderSpecificChips();
    }}, 'x'));
    wrap.appendChild(chip);
  });
}

var searchTimer = null;
function searchOrgs(term) {
  if (searchTimer) clearTimeout(searchTimer);
  var box = document.getElementById('crOrgSearchResults');
  if (!box) return;
  if (!term || term.length < 2) { box.innerHTML = ''; return; }
  searchTimer = setTimeout(function() {
    ajax('search_involvements', {cr_term: term}, function(err, resp) {
      if (err || !resp.success) { box.innerHTML = ''; return; }
      box.innerHTML = '';
      var list = el('div', {class: 'cr-search-results'});
      resp.results.forEach(function(o) {
        var sub = o.program + (o.division ? ' / ' + o.division : '');
        var row = el('div', {class: 'cr-search-result', onclick: function() {
          if (!state.specificOrgs.some(function(x) { return x.orgId === o.orgId; })) {
            state.specificOrgs.push(o);
            renderSpecificChips();
          }
          box.innerHTML = '';
        }});
        row.appendChild(el('div', null, o.orgName));
        if (sub.trim()) row.appendChild(el('div', {class: 'cr-muted'}, sub));
        list.appendChild(row);
      });
      box.appendChild(list);
    });
  }, 220);
}

function bulkAddFromString(s, onDone) {
  ajax('bulk_lookup_involvements', {cr_ids_raw: s}, function(err, resp) {
    if (err || !resp.success) { if (onDone) onDone(); return; }
    (resp.found || []).forEach(function(o) {
      if (!state.specificOrgs.some(function(x) { return x.orgId === o.orgId; })) {
        state.specificOrgs.push(o);
      }
    });
    renderSpecificChips();
    if (resp.missing && resp.missing.length) {
      console.log('IDs not found:', resp.missing);
    }
    if (resp.inactive && resp.inactive.length) {
      console.log('Inactive orgs skipped:', resp.inactive);
    }
    if (onDone) onDone();
  });
}

function runReport() {
  if (!hasAnyScope()) {
    alert('Pick at least one Program/Division or specific Involvement first.');
    return;
  }
  var btn = event && event.target;
  if (btn) { btn.disabled = true; btn.textContent = 'Running...'; }
  ajax('run_report', {
    cr_prog_div_groups: JSON.stringify(state.progDivGroups),
    cr_specific_org_ids: JSON.stringify(state.specificOrgs.map(function(o) { return o.orgId; })),
    cr_date_field: state.dateField,
    cr_date_from: state.dateFrom,
    cr_date_to: state.dateTo,
    cr_status: state.status,
    cr_code_prefix: state.codePrefix,
    cr_name_search: state.nameSearch
  }, function(err, resp) {
    if (btn) { btn.disabled = false; btn.textContent = 'Run Report'; }
    if (err || !resp.success) {
      alert('Error: ' + (err ? err.message : (resp && resp.message) || 'Unknown'));
      return;
    }
    state.result = resp;
    state.view = 'results';
    state.resultTab = 'byName';
    render();
  });
}

function hasAnyScope() {
  for (var i = 0; i < state.progDivGroups.length; i++) {
    var g = state.progDivGroups[i];
    if (g.programId || g.divisionId) return true;
  }
  return state.specificOrgs.length > 0;
}

function renderResults() {
  var d = state.result;
  var card = el('div', {class: 'cr-card'});

  var back = el('span', {class: 'cr-back', onclick: function() {
    state.view = 'scope';
    render();
  }}, '< Back to filters');
  card.appendChild(back);

  card.appendChild(el('div', {class: 'cr-h2'}, 'Results'));

  // Filter summary line
  var fp = [];
  fp.push(state.status === 'all' ? 'All statuses' : ('Status: ' + state.status));
  if (state.dateFrom) fp.push((state.dateField === 'used' ? 'Used' : 'Created') + ' ' + state.dateFrom + ' to ' + (state.dateTo || 'today'));
  if (state.codePrefix) fp.push('Code starts with "' + state.codePrefix + '"');
  if (state.nameSearch) fp.push('Name contains "' + state.nameSearch + '"');
  card.appendChild(el('div', {class: 'cr-sub'}, fp.join(' \\u00B7 ') + ' \\u00B7 ' + (d.summary.scopeOrgCount || 0) + ' involvement(s) in scope'));

  // Summary tiles
  var tiles = el('div', {class: 'cr-tile-row'});
  function tile(num, label, cls) {
    var t = el('div', {class: 'cr-tile ' + (cls || '')});
    t.appendChild(el('div', {class: 'cr-tile-num'}, String(num)));
    t.appendChild(el('div', {class: 'cr-tile-label'}, label));
    return t;
  }
  tiles.appendChild(tile(d.summary.used || 0, 'Used'));
  tiles.appendChild(tile(fmtMoney(d.summary.totalDiscounted), 'Total Discounted', 'cr-tile-money'));
  tiles.appendChild(tile(d.summary.unused || 0, 'Unused', 'cr-tile-unused'));
  tiles.appendChild(tile(d.summary.cancelled || 0, 'Cancelled', 'cr-tile-cancelled'));
  tiles.appendChild(tile(d.summary.total || 0, 'Total Rows'));
  card.appendChild(tiles);

  if (d.truncated) {
    card.appendChild(el('div', {style: 'background:#fff4ce;border:1px solid #f4d35e;padding:8px 12px;border-radius:4px;margin-bottom:10px;font-size:13px;color:#7a5c00;'},
      'Truncated to 5000 rows. Tighten the date range or scope for full output.'));
  }

  // Sub-tabs
  var subtabs = el('div', {class: 'cr-subtabs'});
  ['byName', 'byOrg', 'detail'].forEach(function(t) {
    var label = t === 'byName' ? 'By Name on Code' : (t === 'byOrg' ? 'By Involvement' : 'Details');
    var tab = el('div', {class: 'cr-subtab' + (state.resultTab === t ? ' active' : ''),
      onclick: function() { state.resultTab = t; render(); }}, label);
    subtabs.appendChild(tab);
  });
  // CSV export on the right
  var actions = el('div', {style: 'margin-left:auto; padding:6px;'});
  var csvBtn = el('button', {class: 'cr-btn cr-secondary cr-sm', onclick: exportCsv}, 'Download CSV');
  actions.appendChild(csvBtn);
  subtabs.appendChild(actions);
  card.appendChild(subtabs);

  if (state.resultTab === 'byName') card.appendChild(renderByName(d.byName || []));
  else if (state.resultTab === 'byOrg') card.appendChild(renderByOrg(d.byOrg || []));
  else card.appendChild(renderDetail(d.rows || []));

  root.appendChild(card);
}

function renderByName(list) {
  if (!list.length) return el('div', {class: 'cr-empty'}, 'No coupons in scope.');
  var table = el('table', {class: 'cr-table'});
  var thead = el('thead');
  thead.innerHTML = '<tr><th>Name on Code</th><th class="num">Used</th><th class="num">$ Discounted</th><th class="num">Unused</th><th class="num">Cancelled</th></tr>';
  table.appendChild(thead);
  var tb = el('tbody');
  list.forEach(function(n) {
    var tr = el('tr');
    tr.innerHTML = '<td><strong>' + escapeHtml(n.nameOnCode) + '</strong></td>'
      + '<td class="num">' + (n.used || 0) + '</td>'
      + '<td class="num">' + fmtMoney(n.totalDiscounted) + '</td>'
      + '<td class="num">' + (n.unused || 0) + '</td>'
      + '<td class="num">' + (n.cancelled || 0) + '</td>';
    tb.appendChild(tr);
  });
  table.appendChild(tb);
  return table;
}

function renderByOrg(list) {
  if (!list.length) return el('div', {class: 'cr-empty'}, 'No coupons in scope.');
  var table = el('table', {class: 'cr-table'});
  var thead = el('thead');
  thead.innerHTML = '<tr><th>Involvement</th><th class="num">Used</th><th class="num">$ Discounted</th><th class="num">Unused</th><th class="num">Cancelled</th></tr>';
  table.appendChild(thead);
  var tb = el('tbody');
  list.forEach(function(o) {
    var tr = el('tr');
    var orgLink = '<a href="/Org/' + o.orgId + '" target="_blank">' + escapeHtml(o.orgName) + '</a>';
    tr.innerHTML = '<td>' + orgLink + '</td>'
      + '<td class="num">' + (o.used || 0) + '</td>'
      + '<td class="num">' + fmtMoney(o.totalDiscounted) + '</td>'
      + '<td class="num">' + (o.unused || 0) + '</td>'
      + '<td class="num">' + (o.cancelled || 0) + '</td>';
    tb.appendChild(tr);
  });
  table.appendChild(tb);
  return table;
}

function renderDetail(rows) {
  if (!rows.length) return el('div', {class: 'cr-empty'}, 'No coupons match these filters.');
  var table = el('table', {class: 'cr-table'});
  var thead = el('thead');
  thead.innerHTML = '<tr>'
    + '<th>Status</th>'
    + '<th>Code</th>'
    + '<th>Name on Code</th>'
    + '<th>Person Who Used</th>'
    + '<th>Involvement</th>'
    + '<th class="num">Amount</th>'
    + '<th>Type</th>'
    + '<th>Used</th>'
    + '<th>Created</th>'
    + '<th>Created By</th>'
    + '</tr>';
  table.appendChild(thead);
  var tb = el('tbody');
  rows.forEach(function(r) {
    var statusCls = r.status === 'Used' ? 'cr-status-used'
                  : r.status === 'Cancelled' ? 'cr-status-cancelled'
                  : 'cr-status-unused';
    var typeLabel = r.type || '';
    if (r.type === 'pct' && r.percentage) typeLabel = r.percentage + '%';
    else if (r.type === 'amt') typeLabel = 'fixed';
    var personCell = r.peopleId
        ? '<a href="/Person2/' + r.peopleId + '" target="_blank">' + escapeHtml(r.personName || ('#' + r.peopleId)) + '</a>'
        : '<span class="cr-muted">-</span>';
    var orgCell = r.orgId
        ? '<a href="/Org/' + r.orgId + '" target="_blank">' + escapeHtml(r.orgName) + '</a>'
        : escapeHtml(r.orgName);
    var createdByCell = r.createdById
        ? '<a href="/Person2/' + r.createdById + '" target="_blank">' + escapeHtml(r.createdByName || ('#' + r.createdById)) + '</a>'
        : '<span class="cr-muted">-</span>';
    var tr = el('tr');
    tr.innerHTML = '<td><span class="' + statusCls + '">' + r.status + '</span></td>'
      + '<td style="font-family:monospace; font-size:12px;">' + escapeHtml(r.code) + '</td>'
      + '<td>' + (r.nameOnCode ? '<strong>' + escapeHtml(r.nameOnCode) + '</strong>' : '<span class="cr-muted">(no name)</span>') + '</td>'
      + '<td>' + personCell + '</td>'
      + '<td>' + orgCell + '</td>'
      + '<td class="num">' + fmtMoney(r.amount) + '</td>'
      + '<td>' + escapeHtml(typeLabel) + '</td>'
      + '<td>' + escapeHtml(r.used || '') + '</td>'
      + '<td>' + escapeHtml(r.created || '') + '</td>'
      + '<td>' + createdByCell + '</td>';
    tb.appendChild(tr);
  });
  table.appendChild(tb);
  return table;
}

function exportCsv() {
  if (!state.result || !state.result.rows || !state.result.rows.length) {
    alert('Nothing to export.');
    return;
  }
  var headers = ['Status', 'Code', 'Name on Code', 'Person Who Used', 'PeopleId',
                 'Involvement', 'OrgId', 'Amount', 'Reg Amount', 'Type', 'Percentage',
                 'Used', 'Created', 'Cancelled', 'Created By', 'MultiUse', 'Generated'];
  var lines = [headers.join(',')];

  function esc(v) {
    if (v === null || v === undefined) return '';
    var s = String(v);
    if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\\n') !== -1) {
      s = '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  state.result.rows.forEach(function(r) {
    lines.push([
      esc(r.status), esc(r.code), esc(r.nameOnCode),
      esc(r.personName), esc(r.peopleId || ''),
      esc(r.orgName), esc(r.orgId || ''),
      esc(r.amount || 0), esc(r.regAmount || 0),
      esc(r.type), esc(r.percentage || ''),
      esc(r.used), esc(r.created), esc(r.cancelled),
      esc(r.createdByName), esc(r.multiUse ? 'Y' : 'N'),
      esc(r.generated ? 'Y' : 'N'),
    ].join(','));
  });

  var blob = new Blob([lines.join('\\n')], {type: 'text/csv;charset=utf-8;'});
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = 'coupon-report-' + new Date().toISOString().substring(0, 10) + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(function() { URL.revokeObjectURL(url); }, 1000);
}

// ===== Bootstrap =====
ajax('get_filters', {}, function(err, resp) {
  if (err || !resp.success) {
    root.innerHTML = '<div class="cr-card"><strong style="color:#c0392b;">Failed to load filters.</strong> Refresh and try again.</div>';
    return;
  }
  state.programs = resp.programs || [];
  state.divisions = resp.divisions || [];
  render();
});
</script>
"""
