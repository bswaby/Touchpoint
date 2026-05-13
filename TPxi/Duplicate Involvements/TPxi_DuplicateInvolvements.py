#roles=Edit
#----------------------------------------------------------------------
# TPxi_DuplicateInvolvements.py
#
# Find people who are in more than one involvement within a chosen
# scope, then clean them down to one with Drop / Move actions.
#
#----------------------------------------------------------------------
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
#
# Scope is the UNION of:
#   - One or more Program / Division rows (any program / any division)
#   - Zero or more specific involvements added by search
#
# Member types can be excluded from the duplicate check (e.g., exclude
# Leader types so a leader serving in three classes is not flagged).
#
# Every action is logged with who, what, when, source org, target org.
# Log is stored in TouchPoint content storage as JSON, capped at the
# last N entries.  A "View Action Log" link in the header opens it.
# Each drop/move also writes a person-note on the affected individual.
#
# Storage Keys:
#   DuplicateInv_ActionLog  - JSON list of recent drop/move actions



import json
import datetime
import traceback

model.Header = 'Duplicate Involvements'

APP_VERSION = '1.0.0'

CONFIG = {
    'max_results': 1000,
    'log_max_entries': 1000,
    'log_storage_key': 'DuplicateInv_ActionLog',
}


# ===== Helpers =====

def safe_str(v):
    if v is None:
        return ''
    try:
        return str(v)
    except:
        try:
            return v.encode('utf-8', 'ignore')
        except:
            return ''


def safe_int(v, default=0):
    try:
        return int(v)
    except:
        return default


def get_data(name, default=''):
    try:
        if hasattr(model.Data, name):
            v = getattr(model.Data, name)
            return v if v is not None else default
    except:
        pass
    return default


def now_iso():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def get_current_user():
    """Return (userId, userName) for the logged-in user."""
    try:
        uid = model.UserPeopleId()
        if uid:
            row = q.QuerySql("SELECT Name FROM People WHERE PeopleId = " + str(int(uid)))
            uname = safe_str(row[0].Name) if row else 'Unknown'
            return (uid, uname)
    except:
        pass
    return (0, 'Unknown')


# ===== Log =====

def load_log():
    try:
        text = model.TextContent(CONFIG['log_storage_key'])
        if text:
            return json.loads(text)
    except:
        pass
    return []


def save_log(log):
    try:
        if len(log) > CONFIG['log_max_entries']:
            log = log[-CONFIG['log_max_entries']:]
        model.WriteContentText(CONFIG['log_storage_key'], json.dumps(log), '')
    except:
        pass


def log_action(action, person_id, person_name, from_org_id, from_org_name, to_org_id=0, to_org_name=''):
    uid, uname = get_current_user()
    log = load_log()
    log.append({
        'ts': now_iso(),
        'userId': uid,
        'userName': uname,
        'action': action,
        'peopleId': person_id,
        'peopleName': person_name,
        'fromOrgId': from_org_id,
        'fromOrgName': from_org_name,
        'toOrgId': to_org_id,
        'toOrgName': to_org_name,
    })
    save_log(log)
    return uname


# ===== SQL =====

def resolve_scope_orgs(prog_div_groups, specific_org_ids):
    """Return DISTINCT list of OrganizationIds (active only) in scope."""
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
        sql = """
            SELECT DISTINCT o.OrganizationId
            FROM Organizations o
            JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId
            WHERE """ + " AND ".join(wh)
        try:
            for r in q.QuerySql(sql):
                org_ids.add(r.OrganizationId)
        except:
            pass

    return list(org_ids)


# ===== AJAX Handlers =====

def handle_get_filters():
    progs = []
    for r in q.QuerySql("SELECT Id, Name FROM Program WHERE Name IS NOT NULL ORDER BY Name"):
        progs.append({'id': r.Id, 'name': safe_str(r.Name)})
    divs = []
    for r in q.QuerySql("SELECT Id, Name, ProgId FROM Division WHERE Name IS NOT NULL ORDER BY Name"):
        divs.append({'id': r.Id, 'name': safe_str(r.Name), 'programId': r.ProgId})
    mtypes = []
    for r in q.QuerySql("SELECT Id, Description FROM lookup.MemberType WHERE Description IS NOT NULL ORDER BY Description"):
        mtypes.append({'id': r.Id, 'description': safe_str(r.Description)})
    print json.dumps({
        'success': True,
        'programs': progs,
        'divisions': divs,
        'memberTypes': mtypes,
    })


def handle_search_involvements():
    term = get_data('dup_term', '').strip()
    if not term:
        print json.dumps({'success': True, 'results': []})
        return
    safe_term = term.replace("'", "''")
    sql = """
        SELECT TOP 30
            o.OrganizationId, o.OrganizationName,
            ISNULL(MAX(os.Division), '') AS Division,
            ISNULL(MAX(os.Program), '') AS Program
        FROM Organizations o
        LEFT JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId
        WHERE o.OrganizationStatusId = 30
          AND o.OrganizationName LIKE '%""" + safe_term + """%'
        GROUP BY o.OrganizationId, o.OrganizationName
        ORDER BY o.OrganizationName
    """
    results = []
    for r in q.QuerySql(sql):
        results.append({
            'orgId': r.OrganizationId,
            'orgName': safe_str(r.OrganizationName),
            'division': safe_str(r.Division),
            'program': safe_str(r.Program),
        })
    print json.dumps({'success': True, 'results': results})


def handle_run_search():
    try:
        prog_div_groups = json.loads(get_data('dup_prog_div_groups', '[]'))
    except:
        prog_div_groups = []
    try:
        specific_org_ids = json.loads(get_data('dup_specific_org_ids', '[]'))
    except:
        specific_org_ids = []
    try:
        excluded_mt = json.loads(get_data('dup_excluded_member_types', '[]'))
        excluded_mt = [int(x) for x in excluded_mt]
    except:
        excluded_mt = []

    scope_orgs = resolve_scope_orgs(prog_div_groups, specific_org_ids)

    if not scope_orgs:
        print json.dumps({
            'success': True, 'results': [], 'scopeOrgCount': 0,
            'message': 'Empty scope. Add a program, division, or specific involvement.'
        })
        return

    org_csv = ','.join(str(x) for x in scope_orgs)
    mt_filter = ''
    if excluded_mt:
        mt_csv = ','.join(str(x) for x in excluded_mt)
        mt_filter = " AND om.MemberTypeId NOT IN (" + mt_csv + ")"

    sql = """
        SELECT om.PeopleId,
               p.Name,
               p.Age,
               COUNT(DISTINCT om.OrganizationId) AS InvCount
        FROM OrganizationMembers om
        JOIN People p ON om.PeopleId = p.PeopleId
        WHERE om.OrganizationId IN (""" + org_csv + """)
          AND om.InactiveDate IS NULL
          """ + mt_filter + """
          AND p.IsDeceased = 0
          AND p.ArchivedFlag = 0
        GROUP BY om.PeopleId, p.Name, p.Age
        HAVING COUNT(DISTINCT om.OrganizationId) > 1
        ORDER BY COUNT(DISTINCT om.OrganizationId) DESC, p.Name
    """

    results = []
    for r in q.QuerySql(sql):
        results.append({
            'peopleId': r.PeopleId,
            'name': safe_str(r.Name),
            'age': r.Age if r.Age is not None else None,
            'invCount': r.InvCount,
        })
        if len(results) >= CONFIG['max_results']:
            break

    print json.dumps({
        'success': True,
        'results': results,
        'scopeOrgCount': len(scope_orgs),
    })


def handle_get_person_detail():
    people_id = safe_int(get_data('dup_people_id', 0))
    if people_id <= 0:
        print json.dumps({'success': False, 'message': 'No person ID'})
        return

    try:
        prog_div_groups = json.loads(get_data('dup_prog_div_groups', '[]'))
    except:
        prog_div_groups = []
    try:
        specific_org_ids = json.loads(get_data('dup_specific_org_ids', '[]'))
    except:
        specific_org_ids = []
    try:
        excluded_mt = json.loads(get_data('dup_excluded_member_types', '[]'))
        excluded_mt = [int(x) for x in excluded_mt]
    except:
        excluded_mt = []

    scope_orgs = resolve_scope_orgs(prog_div_groups, specific_org_ids)
    if not scope_orgs:
        print json.dumps({'success': False, 'message': 'Empty scope'})
        return

    org_csv = ','.join(str(x) for x in scope_orgs)
    mt_filter = ''
    if excluded_mt:
        mt_csv = ','.join(str(x) for x in excluded_mt)
        mt_filter = " AND om.MemberTypeId NOT IN (" + mt_csv + ")"

    person_name = 'Unknown'
    person_age = None
    try:
        for pr in q.QuerySql("SELECT ISNULL(Name, '') AS Name, Age FROM People WHERE PeopleId = " + str(people_id)):
            person_name = safe_str(pr.Name)
            try:
                person_age = int(pr.Age) if pr.Age is not None else None
            except:
                person_age = None
            break
    except:
        pass

    sql = """
        SELECT om.OrganizationId AS OrganizationId,
               ISNULL(o.OrganizationName, '') AS OrganizationName,
               CAST(ISNULL(om.MemberTypeId, 0) AS INT) AS MemberTypeId,
               ISNULL(mt.Description, '?') AS MemberType,
               ISNULL(CONVERT(VARCHAR(10), om.EnrollmentDate, 23), '') AS EnrollmentDate,
               ISNULL(MAX(CAST(os.Division AS NVARCHAR(200))), '') AS Division,
               ISNULL(MAX(CAST(os.Program AS NVARCHAR(200))), '') AS Program
        FROM OrganizationMembers om
        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
        LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
        LEFT JOIN OrganizationStructure os ON os.OrgId = o.OrganizationId
        WHERE om.PeopleId = """ + str(people_id) + """
          AND om.OrganizationId IN (""" + org_csv + """)
          AND om.InactiveDate IS NULL
          """ + mt_filter + """
        GROUP BY om.OrganizationId, o.OrganizationName, om.MemberTypeId, mt.Description, om.EnrollmentDate
        ORDER BY o.OrganizationName
    """

    involvements = []
    try:
        for r in q.QuerySql(sql):
            try:
                enroll = safe_str(r.EnrollmentDate)

                org_id = 0
                try:
                    org_id = int(r.OrganizationId)
                except:
                    org_id = 0
                mt_id = 0
                try:
                    mt_id = int(r.MemberTypeId)
                except:
                    mt_id = 0

                involvements.append({
                    'orgId': org_id,
                    'orgName': safe_str(r.OrganizationName),
                    'memberTypeId': mt_id,
                    'memberType': safe_str(r.MemberType),
                    'enrollmentDate': enroll,
                    'division': safe_str(r.Division),
                    'program': safe_str(r.Program),
                })
            except Exception as inner_e:
                # Skip a bad row rather than fail the whole call
                try:
                    model.DebugPrint('DupInv detail row skip: ' + safe_str(inner_e))
                except:
                    pass
    except Exception as outer_e:
        print json.dumps({
            'success': False,
            'message': 'Detail query failed: ' + safe_str(outer_e),
            'trace': traceback.format_exc(),
        })
        return

    print json.dumps({
        'success': True,
        'peopleId': people_id,
        'name': person_name,
        'age': person_age,
        'involvements': involvements,
    })


def _org_name(org_id):
    try:
        rows = q.QuerySql("SELECT OrganizationName FROM Organizations WHERE OrganizationId = " + str(int(org_id)))
        if rows:
            return safe_str(rows[0].OrganizationName)
    except:
        pass
    return ''


def handle_drop():
    people_id = safe_int(get_data('dup_people_id', 0))
    org_id = safe_int(get_data('dup_org_id', 0))
    if people_id <= 0 or org_id <= 0:
        print json.dumps({'success': False, 'message': 'Missing person or org ID'})
        return
    try:
        person = model.GetPerson(people_id)
        if not person:
            print json.dumps({'success': False, 'message': 'Person not found'})
            return
        person_name = safe_str(person.Name)
        org_name = _org_name(org_id)

        model.DropOrgMember(people_id, org_id)

        uname = log_action('drop', people_id, person_name, org_id, org_name)

        note = 'Duplicate Involvements: Removed from "' + org_name + '". By ' + uname + ' on ' + now_iso() + '.'
        try:
            model.AddNote(people_id, note)
        except:
            pass

        print json.dumps({'success': True, 'message': 'Removed from ' + org_name})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Drop failed: ' + safe_str(e)})


def handle_move():
    people_id = safe_int(get_data('dup_people_id', 0))
    from_org_id = safe_int(get_data('dup_from_org_id', 0))
    to_org_id = safe_int(get_data('dup_to_org_id', 0))
    if people_id <= 0 or from_org_id <= 0 or to_org_id <= 0:
        print json.dumps({'success': False, 'message': 'Missing IDs'})
        return
    if from_org_id == to_org_id:
        print json.dumps({'success': False, 'message': 'Source and target are the same'})
        return
    try:
        person = model.GetPerson(people_id)
        if not person:
            print json.dumps({'success': False, 'message': 'Person not found'})
            return
        person_name = safe_str(person.Name)
        from_name = _org_name(from_org_id)
        to_name = _org_name(to_org_id)

        model.MoveToOrg(people_id, from_org_id, to_org_id)

        uname = log_action('move', people_id, person_name, from_org_id, from_name, to_org_id, to_name)

        note = ('Duplicate Involvements: Moved from "' + from_name + '" to "' + to_name +
                '". By ' + uname + ' on ' + now_iso() + '.')
        try:
            model.AddNote(people_id, note)
        except:
            pass

        print json.dumps({'success': True, 'message': 'Moved to ' + to_name})
    except Exception as e:
        print json.dumps({'success': False, 'message': 'Move failed: ' + safe_str(e)})


def handle_get_log():
    limit = safe_int(get_data('dup_limit', 200))
    log = load_log()
    log_recent = list(reversed(log))[:limit]
    print json.dumps({'success': True, 'log': log_recent, 'totalCount': len(log)})


# ===== Dispatch =====

if model.HttpMethod == "post":
    action = get_data('dup_action', '')
    try:
        if action == 'get_filters':
            handle_get_filters()
        elif action == 'search_involvements':
            handle_search_involvements()
        elif action == 'run_search':
            handle_run_search()
        elif action == 'get_person_detail':
            handle_get_person_detail()
        elif action == 'drop':
            handle_drop()
        elif action == 'move':
            handle_move()
        elif action == 'get_log':
            handle_get_log()
        else:
            print json.dumps({'success': False, 'message': 'Unknown action: ' + safe_str(action)})
    except Exception as e:
        print json.dumps({
            'success': False,
            'message': 'Server error: ' + safe_str(e),
            'trace': traceback.format_exc(),
        })
else:
    # ===== HTML / CSS / JS =====
    # Render via model.Form so we're served by /PyScriptForm/...
    # POST handlers above use print json.dumps(...) for raw AJAX responses.
    model.Form = """
<style>
.dup-root { font-family: 'Segoe UI', Arial, sans-serif; color: #222; max-width: 1400px; margin: 0 auto; padding: 12px; }
.dup-card { background: #fff; border: 1px solid #e1e4e8; border-radius: 8px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
.dup-h1 { font-size: 22px; font-weight: 700; margin: 0 0 4px 0; color: #1f4e79; }
.dup-h2 { font-size: 17px; font-weight: 700; margin: 0 0 8px 0; color: #1f4e79; }
.dup-sub { font-size: 13px; color: #555; margin-bottom: 8px; }
.dup-label { display: block; font-size: 12px; color: #555; margin: 8px 0 4px 0; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
.dup-muted { color: #888; font-size: 12px; }
.dup-row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
.dup-input, .dup-select { padding: 7px 10px; font-size: 14px; border: 1px solid #c8ccd0; border-radius: 4px; background: #fff; min-width: 200px; }
.dup-btn { display: inline-block; padding: 8px 14px; font-size: 14px; font-weight: 600; border: 1px solid #1f4e79; background: #1f4e79; color: #fff; border-radius: 4px; cursor: pointer; text-decoration: none; }
.dup-btn:hover { background: #2a5e8e; }
.dup-btn.dup-secondary { background: #fff; color: #1f4e79; }
.dup-btn.dup-secondary:hover { background: #f0f4f8; }
.dup-btn.dup-danger { background: #c0392b; border-color: #c0392b; }
.dup-btn.dup-danger:hover { background: #d04639; }
.dup-btn.dup-success { background: #27ae60; border-color: #27ae60; }
.dup-btn.dup-success:hover { background: #2ecc71; }
.dup-btn.dup-sm { padding: 4px 9px; font-size: 12px; }
.dup-btn:disabled { opacity: 0.5; cursor: not-allowed; }
.dup-pill { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; background: #d6e6f5; color: #1f4e79; }
.dup-progdiv-rows { display: flex; flex-direction: column; gap: 6px; margin-bottom: 8px; }
.dup-progdiv-row { display: flex; gap: 6px; align-items: center; }
.dup-chip { display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px; background: #eef2f7; border: 1px solid #c8d4e3; border-radius: 999px; font-size: 13px; }
.dup-chip-x { cursor: pointer; color: #c0392b; font-weight: 700; padding: 0 2px; }
.dup-chip-row { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 4px; }
.dup-table { width: 100%; border-collapse: collapse; }
.dup-table th { text-align: left; padding: 8px 10px; background: #f0f4f8; font-size: 13px; font-weight: 700; color: #1f4e79; border-bottom: 2px solid #d8e0e8; }
.dup-table td { padding: 8px 10px; border-bottom: 1px solid #eee; font-size: 14px; vertical-align: middle; }
.dup-table tr.dup-row-click { cursor: pointer; }
.dup-table tr.dup-row-click:hover { background: #f9fbfd; }
.dup-toast { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); padding: 10px 18px; border-radius: 6px; color: #fff; font-weight: 600; z-index: 9999; box-shadow: 0 4px 12px rgba(0,0,0,0.2); }
.dup-toast.dup-t-success { background: #27ae60; }
.dup-toast.dup-t-error { background: #c0392b; }
.dup-toast.dup-t-info { background: #1f4e79; }
.dup-mt-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); gap: 4px; max-height: 200px; overflow-y: auto; padding: 8px; border: 1px solid #e1e4e8; border-radius: 4px; background: #fafbfc; }
.dup-mt-item { display: flex; align-items: center; gap: 6px; font-size: 13px; padding: 2px 4px; cursor: pointer; }
.dup-mt-item:hover { background: #fff; border-radius: 3px; }
.dup-log-row { padding: 8px; border-bottom: 1px solid #eee; font-size: 13px; }
.dup-log-row:last-child { border-bottom: none; }
.dup-log-ts { color: #888; font-size: 11px; margin-top: 2px; }
.dup-log-action-drop { background: #fde8e8; padding: 1px 6px; border-radius: 3px; color: #8a2020; font-weight: 600; font-size: 11px; }
.dup-log-action-move { background: #e6f0fc; padding: 1px 6px; border-radius: 3px; color: #1f4e79; font-weight: 600; font-size: 11px; }
.dup-search-results { max-height: 360px; overflow-y: auto; border: 1px solid #e1e4e8; border-radius: 4px; margin-top: 4px; background: #fff; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
.dup-search-wrap .dup-search-results { position: absolute; z-index: 100; width: 100%; }
.dup-search-result { padding: 6px 10px; cursor: pointer; border-bottom: 1px solid #f0f0f0; }
.dup-search-result:hover { background: #f0f4f8; }
.dup-search-result:last-child { border-bottom: none; }
.dup-empty { color: #888; padding: 30px; text-align: center; font-style: italic; }
.dup-back { color: #1f4e79; cursor: pointer; font-weight: 600; font-size: 14px; margin-bottom: 8px; display: inline-block; }
.dup-back:hover { text-decoration: underline; }
.dup-inv-row { display: flex; justify-content: space-between; align-items: center; padding: 10px 12px; border: 1px solid #e1e4e8; border-radius: 6px; margin-bottom: 6px; background: #fff; }
.dup-inv-meta { font-size: 12px; color: #555; margin-top: 2px; }
.dup-modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 9000; display: flex; align-items: center; justify-content: center; }
.dup-modal { background: #fff; border-radius: 8px; max-width: 600px; width: 90%; max-height: 80vh; overflow-y: auto; padding: 20px; position: relative; }
.dup-search-wrap { position: relative; }
</style>

<div class="dup-root" id="dupRoot">
  <div style="display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:8px; gap:12px;">
    <div>
      <div class="dup-h1">Duplicate Involvements</div>
      <div class="dup-sub">Find people in more than one involvement within a scope. Clean down to one.</div>
    </div>
    <div>
      <button class="dup-btn dup-secondary dup-sm" onclick="dupApp.showLog()">View Action Log</button>
    </div>
  </div>
  <div id="dupMain"></div>
</div>

<script>
window.dupApp = (function(){
var state = {
  programs: [],
  divisions: [],
  memberTypes: [],
  progDivGroups: [{programId: 0, divisionId: 0}],
  specificOrgs: [],
  excludedMemberTypes: [],
  results: [],
  scopeOrgCount: 0,
  currentPerson: null,
  view: 'scope',
  log: [],
  logTotal: 0
};

var root = document.getElementById('dupMain');

function el(tag, attrs, children) {
  var e = document.createElement(tag);
  if (attrs) {
    for (var k in attrs) {
      if (k === 'class') e.className = attrs[k];
      else if (k === 'onclick') e.onclick = attrs[k];
      else if (k === 'onchange') e.onchange = attrs[k];
      else if (k === 'oninput') e.oninput = attrs[k];
      else if (k === 'innerHTML') e.innerHTML = attrs[k];
      else if (k === 'style') e.setAttribute('style', attrs[k]);
      else e.setAttribute(k, attrs[k]);
    }
  }
  if (children !== undefined && children !== null) {
    if (!Array.isArray(children)) children = [children];
    children.forEach(function(c){
      if (c == null) return;
      if (typeof c === 'string' || typeof c === 'number') e.appendChild(document.createTextNode(c));
      else e.appendChild(c);
    });
  }
  return e;
}

function ajax(action, data, cb) {
  data = data || {};
  data.dup_action = action;
  var params = [];
  for (var k in data) {
    params.push(encodeURIComponent(k) + '=' + encodeURIComponent(data[k]));
  }
  var xhr = new XMLHttpRequest();
  xhr.open('POST', window.location.pathname + window.location.search, true);
  xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  xhr.onload = function(){
    var txt = xhr.responseText || '';
    try {
      var resp = JSON.parse(txt);
      cb(null, resp);
    } catch(e) {
      cb('Bad response: ' + txt.substring(0, 200), null);
    }
  };
  xhr.onerror = function(){ cb('Network error', null); };
  xhr.send(params.join('&'));
}

function toast(msg, type) {
  var t = el('div', {class: 'dup-toast dup-t-' + (type || 'info')}, msg);
  document.body.appendChild(t);
  setTimeout(function(){ if (t.parentNode) t.parentNode.removeChild(t); }, 3500);
}

function render() {
  root.innerHTML = '';
  if (state.view === 'scope') renderScope();
  else if (state.view === 'results') renderResults();
  else if (state.view === 'detail') renderDetail();
  else if (state.view === 'log') renderLog();
}

function renderScope() {
  var card = el('div', {class: 'dup-card'});
  card.appendChild(el('div', {class: 'dup-h2'}, 'Define your scope'));
  card.appendChild(el('div', {class: 'dup-sub'}, 'Scope is the UNION of every Program/Division row plus any specific involvements you add.'));

  card.appendChild(el('div', {class: 'dup-label'}, 'Programs & Divisions'));
  var pdWrap = el('div', {class: 'dup-progdiv-rows', id: 'dupPdRows'});
  card.appendChild(pdWrap);
  renderProgDivRows();
  var addPdBtn = el('button', {class: 'dup-btn dup-secondary dup-sm', onclick: function(){
    state.progDivGroups.push({programId: 0, divisionId: 0});
    renderProgDivRows();
  }}, '+ Add Program/Division');
  card.appendChild(addPdBtn);

  card.appendChild(el('div', {class: 'dup-label'}, 'Specific Involvements (optional)'));
  var chipRow = el('div', {class: 'dup-chip-row', id: 'dupSpecificChips'});
  card.appendChild(chipRow);
  renderSpecificChips();
  var searchWrap = el('div', {class: 'dup-search-wrap'});
  var searchInput = el('input', {class: 'dup-input', type: 'text', placeholder: 'Search involvements to add...', oninput: function(e){ searchOrgs(e.target.value, 'dupOrgSearchResults', addSpecificOrg); }});
  searchWrap.appendChild(searchInput);
  searchWrap.appendChild(el('div', {id: 'dupOrgSearchResults'}));
  card.appendChild(searchWrap);

  card.appendChild(el('div', {class: 'dup-label'}, 'Exclude these member types from the duplicate check'));
  var mtGrid = el('div', {class: 'dup-mt-grid'});
  state.memberTypes.forEach(function(mt){
    var label = el('label', {class: 'dup-mt-item'});
    var cb = el('input', {type: 'checkbox', value: mt.id});
    cb.checked = state.excludedMemberTypes.indexOf(mt.id) !== -1;
    cb.onchange = function(){
      var id = parseInt(cb.value);
      var idx = state.excludedMemberTypes.indexOf(id);
      if (cb.checked && idx === -1) state.excludedMemberTypes.push(id);
      else if (!cb.checked && idx !== -1) state.excludedMemberTypes.splice(idx, 1);
    };
    label.appendChild(cb);
    label.appendChild(document.createTextNode(' ' + mt.description));
    mtGrid.appendChild(label);
  });
  card.appendChild(mtGrid);
  card.appendChild(el('div', {class: 'dup-muted', style: 'margin-top:4px;'}, 'Tip: exclude Leader, Volunteer, or similar types so people serving across orgs aren\\'t flagged.'));

  card.appendChild(el('div', {style: 'height:14px'}));
  var runBtn = el('button', {class: 'dup-btn', onclick: runSearch}, 'Find Duplicates');
  card.appendChild(runBtn);

  root.appendChild(card);
}

function renderProgDivRows() {
  var wrap = document.getElementById('dupPdRows');
  if (!wrap) return;
  wrap.innerHTML = '';
  state.progDivGroups.forEach(function(g, idx){
    var row = el('div', {class: 'dup-progdiv-row'});
    var pSel = el('select', {class: 'dup-select', onchange: function(e){
      g.programId = parseInt(e.target.value);
      g.divisionId = 0;
      renderProgDivRows();
    }});
    pSel.appendChild(el('option', {value: 0}, '(any program)'));
    state.programs.forEach(function(p){
      var opt = el('option', {value: p.id}, p.name);
      if (p.id === g.programId) opt.selected = true;
      pSel.appendChild(opt);
    });
    var dSel = el('select', {class: 'dup-select', onchange: function(e){
      g.divisionId = parseInt(e.target.value);
    }});
    dSel.appendChild(el('option', {value: 0}, '(any division)'));
    state.divisions.filter(function(d){ return !g.programId || d.programId === g.programId; }).forEach(function(d){
      var opt = el('option', {value: d.id}, d.name);
      if (d.id === g.divisionId) opt.selected = true;
      dSel.appendChild(opt);
    });
    var rm = el('button', {class: 'dup-btn dup-danger dup-sm', onclick: function(){
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
  var wrap = document.getElementById('dupSpecificChips');
  if (!wrap) return;
  wrap.innerHTML = '';
  state.specificOrgs.forEach(function(o, idx){
    var chip = el('span', {class: 'dup-chip'});
    chip.appendChild(document.createTextNode(o.orgName));
    chip.appendChild(el('span', {class: 'dup-chip-x', onclick: function(){
      state.specificOrgs.splice(idx, 1);
      renderSpecificChips();
    }}, 'x'));
    wrap.appendChild(chip);
  });
}

var searchTimer = null;
function searchOrgs(term, resultsBoxId, onPick) {
  if (searchTimer) clearTimeout(searchTimer);
  var box = document.getElementById(resultsBoxId);
  if (!box) return;
  if (!term || term.length < 2) { box.innerHTML = ''; return; }
  searchTimer = setTimeout(function(){
    ajax('search_involvements', {dup_term: term}, function(err, resp){
      if (err || !resp.success) { box.innerHTML = ''; return; }
      box.innerHTML = '';
      var list = el('div', {class: 'dup-search-results'});
      var anyShown = false;
      resp.results.forEach(function(o){
        anyShown = true;
        var sub = o.program + (o.division ? ' / ' + o.division : '');
        var row = el('div', {class: 'dup-search-result', onclick: function(){
          onPick(o);
          box.innerHTML = '';
        }});
        row.appendChild(el('div', null, o.orgName));
        if (sub.trim()) row.appendChild(el('div', {class: 'dup-muted'}, sub));
        list.appendChild(row);
      });
      if (!anyShown) list.appendChild(el('div', {class: 'dup-search-result dup-muted'}, 'No matches'));
      box.appendChild(list);
    });
  }, 250);
}

function addSpecificOrg(o) {
  if (state.specificOrgs.some(function(s){ return s.orgId === o.orgId; })) return;
  state.specificOrgs.push(o);
  renderSpecificChips();
  // Clear the search input
  var inputs = document.querySelectorAll('#dupRoot input[type="text"]');
  for (var i=0; i<inputs.length; i++) inputs[i].value = '';
}

function runSearch() {
  toast('Searching...', 'info');
  ajax('run_search', {
    dup_prog_div_groups: JSON.stringify(state.progDivGroups),
    dup_specific_org_ids: JSON.stringify(state.specificOrgs.map(function(o){ return o.orgId; })),
    dup_excluded_member_types: JSON.stringify(state.excludedMemberTypes)
  }, function(err, resp){
    if (err) { toast(err, 'error'); return; }
    if (!resp.success) { toast(resp.message || 'Search failed', 'error'); return; }
    state.results = resp.results;
    state.scopeOrgCount = resp.scopeOrgCount;
    state.view = 'results';
    render();
  });
}

function renderResults() {
  var card = el('div', {class: 'dup-card'});
  var bar = el('div', {style: 'display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;'});
  bar.appendChild(el('div', null, [
    el('div', {class: 'dup-h2'}, 'Duplicates Found'),
    el('div', {class: 'dup-sub'}, state.results.length + ' people in >1 involvement (scope = ' + state.scopeOrgCount + ' involvements)')
  ]));
  bar.appendChild(el('button', {class: 'dup-btn dup-secondary', onclick: function(){
    state.view = 'scope';
    render();
  }}, 'Edit Scope'));
  card.appendChild(bar);

  if (!state.results.length) {
    card.appendChild(el('div', {class: 'dup-empty'}, 'No duplicates in this scope.'));
  } else {
    var table = el('table', {class: 'dup-table'});
    var thead = el('thead');
    var thr = el('tr');
    ['Name', 'Age', '# Involvements', ''].forEach(function(h){ thr.appendChild(el('th', null, h)); });
    thead.appendChild(thr);
    table.appendChild(thead);
    var tbody = el('tbody');
    state.results.forEach(function(r){
      var row = el('tr', {class: 'dup-row-click', onclick: function(){ openDetail(r.peopleId); }});
      row.appendChild(el('td', null, r.name));
      row.appendChild(el('td', null, r.age == null ? '' : String(r.age)));
      row.appendChild(el('td', null, [el('span', {class: 'dup-pill'}, String(r.invCount))]));
      row.appendChild(el('td', null, [el('button', {class: 'dup-btn dup-secondary dup-sm'}, 'View')]));
      tbody.appendChild(row);
    });
    table.appendChild(tbody);
    card.appendChild(table);
  }
  root.appendChild(card);
}

function openDetail(peopleId) {
  ajax('get_person_detail', {
    dup_people_id: peopleId,
    dup_prog_div_groups: JSON.stringify(state.progDivGroups),
    dup_specific_org_ids: JSON.stringify(state.specificOrgs.map(function(o){ return o.orgId; })),
    dup_excluded_member_types: JSON.stringify(state.excludedMemberTypes)
  }, function(err, resp){
    if (err) { toast(err, 'error'); return; }
    if (!resp.success) {
      var msg = resp.message || 'Load failed';
      if (resp.trace) { console.error('Detail trace:', resp.trace); msg += ' (see console for trace)'; }
      toast(msg, 'error'); return;
    }
    state.currentPerson = resp;
    state.view = 'detail';
    render();
    // Update results count cache
    for (var i = 0; i < state.results.length; i++) {
      if (state.results[i].peopleId === resp.peopleId) {
        state.results[i].invCount = resp.involvements.length;
        if (resp.involvements.length <= 1) {
          state.results.splice(i, 1);
        }
        break;
      }
    }
  });
}

function renderDetail() {
  var p = state.currentPerson;
  var card = el('div', {class: 'dup-card'});
  card.appendChild(el('span', {class: 'dup-back', onclick: function(){
    state.view = 'results';
    state.currentPerson = null;
    render();
  }}, 'Back to results'));

  var bar = el('div', {style: 'display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;'});
  bar.appendChild(el('div', null, [
    el('div', {class: 'dup-h2'}, p.name),
    el('div', {class: 'dup-sub'}, 'Age: ' + (p.age == null ? '?' : p.age) + ' . ' + p.involvements.length + ' involvements in scope')
  ]));
  bar.appendChild(el('a', {href: '/Person2/' + p.peopleId, target: '_blank', class: 'dup-btn dup-secondary dup-sm'}, 'Open Profile'));
  card.appendChild(bar);

  if (!p.involvements.length) {
    card.appendChild(el('div', {class: 'dup-empty'}, 'No involvements remain in scope.'));
  } else {
    p.involvements.forEach(function(inv){
      var row = el('div', {class: 'dup-inv-row'});
      var left = el('div');
      left.appendChild(el('div', null, [
        el('strong', null, inv.orgName),
        ' ',
        el('span', {class: 'dup-pill'}, inv.memberType)
      ]));
      var metaBits = [];
      if (inv.program) metaBits.push(inv.program + (inv.division ? ' / ' + inv.division : ''));
      if (inv.enrollmentDate) metaBits.push('Enrolled ' + inv.enrollmentDate);
      left.appendChild(el('div', {class: 'dup-inv-meta'}, metaBits.join(' . ')));
      row.appendChild(left);

      var actions = el('div', {style: 'display:flex; gap:6px;'});
      actions.appendChild(el('button', {class: 'dup-btn dup-danger dup-sm', onclick: function(){ drop(inv); }}, 'Drop'));
      actions.appendChild(el('button', {class: 'dup-btn dup-success dup-sm', onclick: function(){ openMove(inv); }}, 'Move'));
      row.appendChild(actions);
      card.appendChild(row);
    });
  }
  root.appendChild(card);
}

function drop(inv) {
  if (!confirm('Remove ' + state.currentPerson.name + ' from "' + inv.orgName + '"?')) return;
  ajax('drop', {
    dup_people_id: state.currentPerson.peopleId,
    dup_org_id: inv.orgId
  }, function(err, resp){
    if (err) { toast(err, 'error'); return; }
    if (!resp.success) { toast(resp.message || 'Drop failed', 'error'); return; }
    toast(resp.message, 'success');
    openDetail(state.currentPerson.peopleId);
  });
}

function openMove(inv) {
  var overlay = el('div', {class: 'dup-modal-overlay', onclick: function(e){
    if (e.target === overlay) document.body.removeChild(overlay);
  }});
  var modal = el('div', {class: 'dup-modal'});
  modal.appendChild(el('div', {class: 'dup-h2'}, 'Move from "' + inv.orgName + '"'));
  modal.appendChild(el('div', {class: 'dup-sub'}, 'Pick a target involvement (active orgs only).'));
  var searchInput = el('input', {class: 'dup-input', style: 'width:100%; box-sizing:border-box;', placeholder: 'Search involvements...'});
  searchInput.oninput = function(e){
    searchOrgs(e.target.value, 'dupMoveResults', function(toOrg){
      if (toOrg.orgId === inv.orgId) { toast('Cannot move to the same involvement', 'error'); return; }
      if (!confirm('Move ' + state.currentPerson.name + ' from "' + inv.orgName + '" to "' + toOrg.orgName + '"?')) return;
      ajax('move', {
        dup_people_id: state.currentPerson.peopleId,
        dup_from_org_id: inv.orgId,
        dup_to_org_id: toOrg.orgId
      }, function(err, resp){
        if (err) { toast(err, 'error'); return; }
        if (!resp.success) { toast(resp.message || 'Move failed', 'error'); return; }
        toast(resp.message, 'success');
        document.body.removeChild(overlay);
        openDetail(state.currentPerson.peopleId);
      });
    });
  };
  modal.appendChild(searchInput);
  modal.appendChild(el('div', {id: 'dupMoveResults', style: 'margin-top:8px; position:relative;'}));
  modal.appendChild(el('div', {style: 'margin-top:16px; text-align:right;'}, [
    el('button', {class: 'dup-btn dup-secondary', onclick: function(){ document.body.removeChild(overlay); }}, 'Cancel')
  ]));
  overlay.appendChild(modal);
  document.body.appendChild(overlay);
  setTimeout(function(){ searchInput.focus(); }, 50);
}

function showLog() {
  ajax('get_log', {dup_limit: 200}, function(err, resp){
    if (err) { toast(err, 'error'); return; }
    if (!resp.success) { toast(resp.message || 'Log load failed', 'error'); return; }
    state.log = resp.log;
    state.logTotal = resp.totalCount;
    state.view = 'log';
    render();
  });
}

function renderLog() {
  var card = el('div', {class: 'dup-card'});
  card.appendChild(el('span', {class: 'dup-back', onclick: function(){
    state.view = state.results.length ? 'results' : 'scope';
    render();
  }}, 'Back'));
  card.appendChild(el('div', {class: 'dup-h2'}, 'Action Log'));
  card.appendChild(el('div', {class: 'dup-sub'}, 'Showing ' + state.log.length + ' of ' + state.logTotal + ' recent actions (newest first).'));

  if (!state.log.length) {
    card.appendChild(el('div', {class: 'dup-empty'}, 'No actions logged yet.'));
  } else {
    state.log.forEach(function(e){
      var row = el('div', {class: 'dup-log-row'});
      var pill = el('span', {class: 'dup-log-action-' + e.action}, e.action.toUpperCase());
      var msg;
      if (e.action === 'drop') {
        msg = el('span', null, [' ', el('strong', null, e.userName), ' dropped ', el('strong', null, e.peopleName), ' from "' + e.fromOrgName + '"']);
      } else if (e.action === 'move') {
        msg = el('span', null, [' ', el('strong', null, e.userName), ' moved ', el('strong', null, e.peopleName), ' from "' + e.fromOrgName + '" to "' + e.toOrgName + '"']);
      } else {
        msg = el('span', null, [' ', e.userName, ' ', e.action]);
      }
      row.appendChild(el('div', null, [pill, msg]));
      row.appendChild(el('div', {class: 'dup-log-ts'}, e.ts));
      card.appendChild(row);
    });
  }
  root.appendChild(card);
}

function boot() {
  ajax('get_filters', {}, function(err, resp){
    if (err) { toast('Init failed: ' + err, 'error'); return; }
    if (!resp.success) { toast('Init failed', 'error'); return; }
    state.programs = resp.programs;
    state.divisions = resp.divisions;
    state.memberTypes = resp.memberTypes;
    state.view = 'scope';
    render();
  });
}

boot();

return {
  showLog: showLog
};
})();
</script>
"""
