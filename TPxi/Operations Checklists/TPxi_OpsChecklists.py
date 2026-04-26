#roles=Access
#--------------------------------------------------------------------
# TPxi Operations Checklists v1.2.4
# Group-based recurring operations management
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
#
# Access model:
#   #roles=Access lets any logged-in user open the page. What they can DO
#   is enforced inside the script via the Settings panel:
#     - Edit Roles: who can add/edit/delete groups, checks, and library items
#     - Complete Roles: who can mark checks complete, toggle steps, and run
#   Admins always have full access. To restrict the page itself further,
#   change #roles=Access above to a specific role (e.g. #roles=Staff).
#--------------------------------------------------------------------
import json
import datetime

# --- Version / Auto-update -------------------------------------------
APP_VERSION = '1.2.4'
DC_SCRIPT_ID = 'TPxi_OpsChecklists'  # ID used on DisplayCache to identify this script
# Use workers.dev URL for server-side fetches (bypasses Cloudflare Bot Fight Mode)
# scripts.displaycache.com is the custom domain used for browser-side version checks.
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'

model.Header = "Operations Checklists"

# --- Helpers ---------------------------------------------------------
def safe(v):
    if v is None:
        return ''
    try:
        return str(v).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;').replace("'",'&#39;')
    except:
        return ''

GK = "TPxi_OpsChecklist_Groups"
SK = "TPxi_OpsChecklist_Steps"
LK = "TPxi_OpsChecklist_Log"
RK = "TPxi_OpsChecklist_Results"
STK = "TPxi_OpsChecklist_Settings"

CATEGORIES = ["People","Involvements","Email","Finance","Facilities","General","Archive"]

def gck(gid):
    return "TPxi_OpsChecklist_G_" + str(gid)

def lj(n, d="{}"):
    try:
        r = model.TextContent(n)
        if r:
            return json.loads(r)
    except:
        pass
    try:
        return json.loads(d)
    except:
        return {}

def _clean_for_json(o):
    """Recursively walk a value and convert any byte strings to unicode so
    json.dumps doesn't choke on Latin-1/CP1252 bytes (Spanish names with
    n-tilde, accented chars, etc.). Idempotent and safe on already-clean data."""
    try:
        if o is None or isinstance(o, (int, float, bool)):
            return o
    except:
        pass
    try:
        if isinstance(o, unicode):
            return o
    except NameError:
        pass
    if isinstance(o, dict):
        cleaned = {}
        for k, v in o.items():
            ck = _clean_for_json(k) if not isinstance(k, str) else (k.decode('utf-8', errors='replace') if hasattr(k, 'decode') else k)
            cleaned[ck] = _clean_for_json(v)
        return cleaned
    if isinstance(o, (list, tuple)):
        return [_clean_for_json(x) for x in o]
    # Bytes / str — decode safely
    try:
        if hasattr(o, 'decode'):
            try:
                return o.decode('utf-8')
            except:
                try:
                    return o.decode('latin-1')
                except:
                    return o.decode('latin-1', errors='replace')
    except:
        pass
    try:
        return unicode(o)
    except:
        try:
            return str(o).decode('latin-1', errors='replace')
        except:
            return ''

def sj(n, d):
    # Layered fallback: try direct write; if json.dumps chokes on bytes,
    # deep-clean and retry; if THAT still fails, drop 'detail' rows (most
    # likely culprit since they come from People objects) and retry.
    # Bare except since IronPython may surface CLR exceptions that don't
    # subclass Python's UnicodeDecodeError.
    try:
        model.WriteContentText(n, json.dumps(d), "")
        return
    except:
        pass
    try:
        model.WriteContentText(n, json.dumps(_clean_for_json(d)), "")
        return
    except:
        pass
    # Last resort: strip 'detail' (the rows from QueryList/SQL) which is
    # the most likely source of byte data we can't safely encode.
    try:
        stripped = d
        if isinstance(d, dict):
            stripped = {}
            for k, v in d.items():
                if isinstance(v, dict) and 'detail' in v:
                    nv = dict(v)
                    nv.pop('detail', None)
                    nv.pop('columns', None)
                    nv['_detailDropped'] = True
                    stripped[k] = nv
                else:
                    stripped[k] = v
        model.WriteContentText(n, json.dumps(_clean_for_json(stripped)), "")
    except Exception as _sj_e:
        # Absolute last resort — log to debug, don't crash the request
        try:
            model.DebugPrint("sj() final failure for " + str(n) + ": " + str(_sj_e))
        except:
            pass

def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def today():
    return datetime.datetime.now().strftime("%Y-%m-%d")

_uid_counter = [0]
def uid():
    """Unique ID with microsecond precision + an in-process counter so even
    requests at the same instant don't collide. Prior versions used only
    second-precision which could create duplicate check IDs and cause
    delete-one-remove-many bugs."""
    _uid_counter[0] = (_uid_counter[0] + 1) % 100000
    return datetime.datetime.now().strftime("%Y%m%d%H%M%S%f") + str(_uid_counter[0])

def user():
    try:
        return str(model.UserName)
    except:
        return 'Unknown'

def get_script_name():
    """Detect the actual script name. Order of preference:
      1. script_name posted from JS (most reliable; window.location.pathname)
      2. Parsed from model.URL (may not work in all IronPython contexts)
      3. Previously persisted scriptName in settings (covers batch context)
      4. Hardcoded default
    Lets the email link and wrapper-script instructions stay correct even if
    an admin renames the script."""
    try:
        if hasattr(model, 'Data') and hasattr(model.Data, 'script_name'):
            sn_in = str(model.Data.script_name).strip()
            if sn_in:
                return sn_in
    except:
        pass
    try:
        url = str(getattr(model, 'URL', '') or '')
        import re
        m = re.search(r'/PyScript(?:Form)?/([^/?#&]+)', url)
        if m:
            return m.group(1)
    except:
        pass
    try:
        settings = lj(STK, "{}")
        nm = settings.get("scriptName", "")
        if nm:
            return str(nm)
    except:
        pass
    return "TPxi_OpsChecklists"

def run_sql_to_rows(sql, max_rows=200):
    """Execute SQL and return rows + column metadata for grid display.
    Backward-compatible: if the result is a single row with a 'cnt' column,
    treats it as a count-only query (no detail rows).
    Returns dict with: rows (list of dicts), columns (ordered list of names),
    count (int, -1 on error), error (str), truncated (bool)."""
    import re
    columns = []
    sql_columns = []
    # Parse column order from the SELECT clause for predictable display order
    try:
        parse_sql = sql.strip()
        if re.match(r'\s*WITH\s', parse_sql, re.IGNORECASE):
            all_selects = [m.start() for m in re.finditer(r'\bSELECT\b', parse_sql, re.IGNORECASE)]
            select_match = None
            for pos in reversed(all_selects):
                chunk = parse_sql[pos:]
                m = re.match(r'SELECT\s+(TOP\s+\d+\s+)?(.*?)\s+FROM\s', chunk, re.IGNORECASE | re.DOTALL)
                if m:
                    select_match = m
                    break
        else:
            select_match = re.search(r'SELECT\s+(TOP\s+\d+\s+)?(.*?)\s+FROM\s', parse_sql, re.IGNORECASE | re.DOTALL)
        if select_match:
            for chunk in select_match.group(2).split(','):
                chunk = chunk.strip()
                if not chunk:
                    continue
                as_match = re.search(r'\bAS\s+\[?(\w+)\]?\s*$', chunk, re.IGNORECASE)
                if as_match:
                    sql_columns.append(as_match.group(1))
                else:
                    word_match = re.search(r'(\w+)\s*$', chunk)
                    if word_match:
                        sql_columns.append(word_match.group(1))
    except:
        pass

    rows_out = []
    truncated_flag = False
    total_seen = 0
    try:
        first = True
        _SKIP_ATTRS = set(['Count','Keys','Values','Items','GetType','ToString','Equals','GetHashCode','ReferenceEquals'])
        for r in q.QuerySql(sql):
            total_seen += 1
            if first:
                available = set()
                for attr in dir(r):
                    if attr.startswith('_') or attr in _SKIP_ATTRS:
                        continue
                    try:
                        val = getattr(r, attr)
                        if callable(val):
                            continue
                        available.add(attr)
                    except:
                        pass
                if sql_columns:
                    avail_lower = {a.lower(): a for a in available}
                    for sc in sql_columns:
                        actual = avail_lower.get(sc.lower())
                        if actual and actual not in columns:
                            columns.append(actual)
                    for a in sorted(available):
                        if a not in columns:
                            columns.append(a)
                else:
                    columns = sorted(list(available))
                first = False
            if len(rows_out) >= max_rows:
                truncated_flag = True
                continue
            row = {}
            for col in columns:
                try:
                    v = getattr(r, col)
                    row[col] = _safe_str(v)
                except:
                    row[col] = ''
            rows_out.append(row)
    except Exception as e:
        return {"rows": [], "columns": [], "count": -1, "error": str(e), "truncated": False}

    count = total_seen
    # Backward-compat: single row with 'cnt' column = count-only query
    if not truncated_flag and total_seen == 1 and 'cnt' in columns:
        try:
            count = int(rows_out[0]['cnt'])
            rows_out = []
            columns = []
        except:
            pass
    return {"rows": rows_out, "columns": columns, "count": count, "error": "", "truncated": truncated_flag}

def _safe_str(v):
    """Return a unicode string safely. TouchPoint sometimes returns name
    fields as Latin-1/CP1252 byte strings (Spanish names with ñ, accented
    chars, etc.) that json.dumps fails to encode as UTF-8. This function
    tries UTF-8 first, falls back to Latin-1 with replacement."""
    if v is None:
        return ''
    # Already a unicode string — return as-is
    try:
        if isinstance(v, unicode):
            return v
    except NameError:
        pass  # Python 3 — all strings are unicode
    # If it's a byte string, try UTF-8 then Latin-1
    try:
        if hasattr(v, 'decode'):
            try:
                return v.decode('utf-8')
            except:
                try:
                    return v.decode('latin-1')
                except:
                    return v.decode('latin-1', errors='replace')
    except:
        pass
    # Fallback: convert via str() and decode if needed
    try:
        s = str(v)
        try:
            return s.decode('utf-8') if hasattr(s, 'decode') else s
        except:
            return s.decode('latin-1', errors='replace') if hasattr(s, 'decode') else s
    except:
        return ''

def _normalize_search_code(code):
    """Per TouchPoint docs, multi-condition queries passed to QueryCount /
    QueryList need SPACES around comparison operators:
        MemberStatusId = 10[Member] AND GenderId = 1[Male]   <-- works
        MemberStatusId=10[Member] AND GenderId=1[Male]       <-- often fails
    The Search Builder UI omits the spaces in its 'View Code' output, so we
    add them back here. Bracketed annotations like [Member]/[True]/[False]
    are SUPPORTED per docs (we keep them). Function args like
    RecentVisitCount(Days=14,Prog=0,...) have '=' inside parens — those must
    be left alone, so we use a depth-tracking state machine instead of regex.
    Saved-search names (no comparison operators) are returned unchanged."""
    if not code:
        return ''
    s = str(code)
    if not any(op in s for op in ['=', '<', '>']):
        return s.strip()
    out = []
    i = 0
    n = len(s)
    depth = 0
    while i < n:
        ch = s[i]
        if ch == '(':
            # TP parser also wants a space inside the open paren for function calls:
            # 'IsMemberOf( Prog=1118 )' parses; 'IsMemberOf(Prog=1118)' often doesn't
            depth += 1
            out.append('(')
            if i + 1 < n and s[i+1] != ' ':
                out.append(' ')
            i += 1
        elif ch == ')':
            depth -= 1
            if out and out[-1] != ' ':
                out.append(' ')
            out.append(')')
            i += 1
        elif depth == 0 and ch in '=<>!':
            two = s[i:i+2]
            if two in ('>=', '<=', '!=', '<>'):
                op = '<>' if two == '!=' else two  # != -> <> for compatibility
                if out and out[-1] != ' ':
                    out.append(' ')
                out.append(op)
                if i + 2 < n and s[i+2] != ' ':
                    out.append(' ')
                i += 2
            elif ch in '=<>':
                if out and out[-1] != ' ':
                    out.append(' ')
                out.append(ch)
                if i + 1 < n and s[i+1] != ' ':
                    out.append(' ')
                i += 1
            else:
                out.append(ch); i += 1
        else:
            out.append(ch); i += 1
    import re
    return re.sub(r'\s+', ' ', ''.join(out)).strip()

def run_search_to_rows(search_code, max_rows=200):
    """Execute a Search Builder query (saved name or condition code) and
    return rows of People for grid drilldown. Returns dict with rows,
    columns, count, error, truncated.

    QueryCount supports both saved-search names AND raw condition codes
    in nearly all TP versions. QueryList is more restrictive — some condition
    codes with [True]/[False] annotations or function syntax (RecentVisitCount,
    BirthdayThisWeek, etc.) work in QueryCount but error in QueryList. We
    silently swallow those QueryList errors so the count badge keeps working
    like before; the View button just won't appear (no detail rows)."""
    rows_out = []
    columns = []
    truncated_flag = False
    normalized = _normalize_search_code(search_code)
    try:
        cnt = int(q.QueryCount(normalized))
    except Exception as e:
        # Try raw if normalized failed
        try:
            cnt = int(q.QueryCount(search_code))
        except Exception as e2:
            return {"rows": [], "columns": [], "count": -1, "error": "Search count failed: " + str(e2), "truncated": False}
    try:
        cols = ['PeopleId', 'Name', 'EmailAddress', 'CellPhone', 'MemberStatusId', 'Age']
        people = q.QueryList(normalized)
        for p in people:
            if len(rows_out) >= max_rows:
                truncated_flag = True
                continue
            row = {}
            for col in cols:
                try:
                    v = getattr(p, col, None)
                    row[col] = _safe_str(v)
                except:
                    row[col] = ''
            rows_out.append(row)
        columns = cols
    except:
        # QueryList doesn't accept this code — that's fine, we have the count.
        # Just leave detail rows empty; the View button will hide itself.
        rows_out = []
        columns = []
    return {"rows": rows_out, "columns": columns, "count": cnt, "error": "", "truncated": truncated_flag}

def _send_completion_notice(check, group, user_name, completed_at, notes, pids_csv):
    """Email completion notice to picked PeopleIds. Silent on errors so a
    notification failure never blocks the actual completion. Returns count sent."""
    if not pids_csv:
        return 0
    pids = []
    for p in str(pids_csv).split(','):
        p = p.strip()
        if p:
            try:
                pids.append(int(p))
            except:
                pass
    if not pids:
        return 0
    # From-address: use default recipient's email or fall back
    settings = lj(STK, "{}")
    from_addr = settings.get("defaultEmail", "") or ""
    default_pid = settings.get("defaultRecipientPid", 0)
    if default_pid:
        resolved = _resolve_pid_email(default_pid)
        if resolved:
            from_addr = resolved
    if not from_addr:
        return 0  # Can't send without a from address
    queued_by = pids[0]
    try:
        upid = model.UserPeopleId
        if upid:
            queued_by = upid
    except:
        pass
    cname = check.get("name", "") if check else ""
    gname = group.get("name", "") if group else ""
    gid = group.get("id", "") if group else ""
    subject = "Checklist completed: " + cname
    body = '<div style="font-family:Segoe UI,sans-serif;max-width:600px;margin:0 auto">'
    body += '<div style="background:linear-gradient(135deg,#107c10,#0a5d0a);color:#fff;padding:20px 24px;border-radius:8px 8px 0 0">'
    body += '<h2 style="margin:0;font-size:20px">&#10003; Checklist Completed</h2>'
    body += '<p style="margin:4px 0 0;opacity:.9;font-size:14px">' + safe(cname) + '</p></div>'
    body += '<div style="background:#fff;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;padding:20px;font-size:13px;color:#333">'
    if gname:
        body += '<p style="margin:6px 0"><strong>Group:</strong> ' + safe(gname) + '</p>'
    body += '<p style="margin:6px 0"><strong>Completed by:</strong> ' + safe(user_name) + '</p>'
    body += '<p style="margin:6px 0"><strong>When:</strong> ' + safe(completed_at) + '</p>'
    if notes:
        body += '<p style="margin:14px 0 4px"><strong>Notes:</strong></p>'
        body += '<div style="background:#f8f9fa;padding:10px;border-left:3px solid #107c10;border-radius:0 6px 6px 0;font-size:13px;color:#555">' + safe(notes) + '</div>'
    body += '<div style="margin-top:22px;text-align:center">'
    try:
        host = str(model.CmsHost)
    except:
        host = ''
    sn = get_script_name()
    link = host + '/PyScriptForm/' + sn
    if gid:
        link += '#group=' + str(gid)
    body += '<a href="' + link + '" style="display:inline-block;padding:10px 24px;background:#107c10;color:#fff;text-decoration:none;border-radius:6px;font-size:13px;font-weight:600">Open Checklists</a>'
    body += '</div></div></div>'
    sent = 0
    for pid in pids:
        try:
            model.Email(pid, queued_by, from_addr, "Operations Checklists", subject, body)
            sent += 1
        except:
            pass
    return sent

def _resolve_pid_email(pid):
    """Look up the current email for a PeopleId. Returns '' if person not
    found, archived, deceased, or has no email. Used by send_reminder_emails
    to resolve picker-stored PeopleIds to current email at send time."""
    if not pid:
        return ''
    try:
        sql = ("SELECT TOP 1 EmailAddress, EmailAddress2 FROM People "
               "WHERE PeopleId = " + str(int(pid)) + " "
               "AND IsDeceased = 0 AND ArchivedFlag = 0")
        for row in q.QuerySql(sql):
            em = _safe_str(getattr(row, 'EmailAddress', '')) or _safe_str(getattr(row, 'EmailAddress2', ''))
            return em
    except:
        pass
    return ''

def find_person_id_by_email(email_addr):
    """Look up PeopleId by email address (primary or secondary).
    Skips deceased and archived people. Returns None if no match.
    Uses SQL because model.FindPersonId() requires multiple args in some TP versions."""
    if not email_addr:
        return None
    safe_email = str(email_addr).strip().replace("'", "''")
    if not safe_email:
        return None
    sql = (
        "SELECT TOP 1 PeopleId FROM People "
        "WHERE (EmailAddress = '" + safe_email + "' OR EmailAddress2 = '" + safe_email + "') "
        "AND IsDeceased = 0 AND ArchivedFlag = 0 "
        "ORDER BY CASE WHEN EmailAddress = '" + safe_email + "' THEN 0 ELSE 1 END, PeopleId"
    )
    try:
        for row in q.QuerySql(sql):
            return row.PeopleId
    except:
        pass
    return None

# --- Role / Permission Helpers --------------------------------------
def _user_in_role(role_name):
    if not role_name:
        return False
    try:
        return bool(model.UserIsInRole(role_name))
    except:
        return False

def is_admin():
    return _user_in_role("Admin")

def has_any_role(roles_csv):
    """True if user is Admin or holds at least one role in the csv list."""
    if is_admin():
        return True
    if not roles_csv:
        return False
    for r in str(roles_csv).split(','):
        r = r.strip()
        if r and _user_in_role(r):
            return True
    return False

def can_edit():
    settings = lj(STK, "{}")
    er = settings.get("editRoles", "Admin")
    return has_any_role(er)

def can_complete():
    if can_edit():
        return True
    settings = lj(STK, "{}")
    cr = settings.get("completeRoles", "Admin")
    return has_any_role(cr)

# Action permission classes (used by handle_ajax)
_EDIT_ACTIONS = set([
    'save_group', 'delete_group',
    'save_check', 'delete_check',
    'sync_library', 'update_from_marketplace', 'update_all_marketplace',
    'setup', 'reset_group',
    'submit_to_marketplace', 'lookup_search',
    'bulk_tag', 'dedupe_check_ids'
])
_COMPLETE_ACTIONS = set([
    'complete_check', 'toggle_step', 'save_step_note',
    'run_checks', 'run_single', 'get_check_detail',
    'search_people', 'resolve_people',
    'search_saved_searches'
])
_ADMIN_ACTIONS = set(['save_settings', 'send_reminders_now', 'install_batch', 'uninstall_batch', 'check_batch_install', 'apply_update'])

# Markers used to safely add/remove our block in the MorningBatch script
# without disturbing other scripts that share the morning batch.
_BATCH_MARKER_START = "# >>> TPxi_OpsChecklists batch start (managed by app, do not edit) >>>"
_BATCH_MARKER_END = "# <<< TPxi_OpsChecklists batch end <<<"

# Allowed colors and icons (validated server-side; user-supplied values
# outside these lists are rejected and replaced with safe defaults)
_ALLOWED_ICONS = set(['&#128202;','&#9961;','&#128176;','&#128100;','&#128273;','&#128197;','&#9881;','&#127968;','&#128231;','&#128640;','&#9889;','&#128295;','&#9733;','&#128203;','&#128220;'])
_ALLOWED_COLORS = set(['#0078d4','#107c10','#7a6400','#7c3aed','#d13438','#ea580c','#0891b2','#4f46e5','#be185d','#059669'])
_ALLOWED_CATS = set(["People","Involvements","Email","Finance","Facilities","General","Archive"])
_ALLOWED_FREQS = set(["daily","weekly","monthly","annual"])
_ALLOWED_TYPES = set(["manual","auto","search"])

def safe_icon(v):
    return v if v in _ALLOWED_ICONS else '&#128203;'

def safe_color(v):
    return v if v in _ALLOWED_COLORS else '#0078d4'

def safe_cat(v):
    return v if v in _ALLOWED_CATS else 'General'

def safe_freq(v):
    return v if v in _ALLOWED_FREQS else 'monthly'

def safe_type(v):
    return v if v in _ALLOWED_TYPES else 'manual'

def _make_mp_snapshot(tpl):
    """Snapshot of catalog item fields the user might edit locally.
    Used for drift detection so Update All can warn about/skip edited items
    instead of silently overwriting customizations."""
    return {
        "sql": tpl.get("sql", "") or "",
        "search": tpl.get("search", "") or "",
        "inst": tpl.get("inst", "") or "",
        "steps": list(tpl.get("steps", []) or []),
        "th": tpl.get("th", 0) or 0,
        "cr": tpl.get("cr", 5) or 5,
        "freq": tpl.get("freq", "monthly") or "monthly",
        "cat": tpl.get("cat", "General") or "General",
        "name": tpl.get("name", "") or "",
    }

def _fetch_marketplace_safe():
    """Fetch the marketplace catalog. Cloudflare bot mode commonly blocks
    server-side TouchPoint requests with a 'Just a moment...' JS challenge
    that returns HTML instead of JSON. We try the public endpoint with a
    browser-like User-Agent first; if that's challenged, we fall back to the
    admin endpoint with the configured bearer token (which typically bypasses
    challenges since admin paths run in their own ruleset).
    Returns (items_list, error_message). error_message is '' on success."""

    # Browser-like UA tends to satisfy CF's lighter checks
    browser_ua = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

    def _try_get(url, headers):
        try:
            return model.RestGet(url, headers), None
        except Exception as e:
            return None, "RestGet failed: " + str(e)

    def _looks_like_challenge(raw):
        if not raw:
            return False
        low = str(raw)[:300].lower()
        return ("just a moment" in low) or ("challenge" in low) or ("cf-mitigated" in low)

    def _parse(raw):
        if not raw:
            return None, "empty response"
        s = str(raw).strip()
        if _looks_like_challenge(s):
            snippet = s[:120].replace('\n', ' ').replace('\r', ' ')
            return None, "Cloudflare challenge intercepted (first 120: " + snippet + ")"
        if not s.startswith("{") and not s.startswith("["):
            snippet = s[:120].replace('\n', ' ').replace('\r', ' ')
            return None, "non-JSON (first 120: " + snippet + ")"
        try:
            data = json.loads(s)
            return data, None
        except Exception as e:
            snippet = s[:120].replace('\n', ' ').replace('\r', ' ')
            return None, "JSON parse failed: " + str(e) + " (first 120: " + snippet + ")"

    # Attempt 1: public catalog with browser-like UA
    raw, err = _try_get(
        "https://scripts.displaycache.com/api/ops-catalog",
        {"User-Agent": browser_ua, "Accept": "application/json"}
    )
    if not err:
        data, perr = _parse(raw)
        if data is not None:
            return (data.get("items", []) or [], "")
        # Public endpoint was reachable but returned a challenge — try admin fallback
        public_err = perr or "unknown"
    else:
        public_err = err

    # Attempt 2: admin endpoint via configured bearer token
    settings = lj(STK, "{}")
    bearer = settings.get("marketplaceBearer", "")
    if bearer:
        raw2, err2 = _try_get(
            "https://scripts.displaycache.com/admin/api/ops-catalog",
            {"User-Agent": browser_ua, "Accept": "application/json",
             "Authorization": "Bearer " + bearer}
        )
        if not err2:
            data2, perr2 = _parse(raw2)
            if data2 is not None:
                return (data2.get("items", []) or [], "")
            return ([], "admin fallback returned: " + (perr2 or "unknown"))
        return ([], "admin fallback failed: " + err2 + " | public: " + public_err)

    return ([], "Public endpoint blocked by Cloudflare. Either disable Bot Fight Mode for /api/ops-catalog in your Cloudflare dashboard, OR set 'marketplaceBearer' in TPxi_OpsChecklist_Settings to your admin token. Detail: " + public_err)

def _get_marketplace_items_for_action():
    """Return marketplace items, preferring catalog POSTed from the browser
    in `catalog_json` (which the browser fetched directly — bypasses CF bot
    mode that blocks TouchPoint's Azure egress IPs). Falls back to a
    server-side fetch only if no catalog was supplied. Same return shape as
    _fetch_marketplace_safe: (items, error_message)."""
    try:
        if hasattr(model.Data, 'catalog_json'):
            raw = str(model.Data.catalog_json or '').strip()
            if raw:
                data = json.loads(raw)
                if isinstance(data, dict):
                    items = data.get('items', []) or []
                elif isinstance(data, list):
                    items = data
                else:
                    items = []
                if isinstance(items, list) and len(items) > 0:
                    return (items, '')
    except Exception:
        pass
    return _fetch_marketplace_safe()

def _check_has_local_edits(check):
    """True if a catalog-sourced check has been modified locally compared
    to its stored marketplace snapshot. Returns False for non-catalog items
    or items installed before snapshots existed (no false positives)."""
    snap = check.get("mp_snapshot")
    if not isinstance(snap, dict):
        return False
    if (check.get("sql", "") or "") != (snap.get("sql", "") or ""):
        return True
    if (check.get("search", "") or "") != (snap.get("search", "") or ""):
        return True
    if (check.get("inst", "") or "") != (snap.get("inst", "") or ""):
        return True
    if (check.get("th", 0) or 0) != (snap.get("th", 0) or 0):
        return True
    if (check.get("cr", 5) or 5) != (snap.get("cr", 5) or 5):
        return True
    if (check.get("freq", "") or "") != (snap.get("freq", "") or ""):
        return True
    if (check.get("cat", "") or "") != (snap.get("cat", "") or ""):
        return True
    if (check.get("name", "") or "") != (snap.get("name", "") or ""):
        return True
    if list(check.get("steps", []) or []) != list(snap.get("steps", []) or []):
        return True
    return False

def period_key(freq, dt_str=None, due_day=None, due_month=None):
    """Return a string key for the current period of a frequency.
    due_day: for weekly 0-6 (Mon-Sun), for monthly 1-28, for annual 1-31
    due_month: for annual 1-12"""
    if dt_str:
        try:
            dt = datetime.datetime(int(dt_str[:4]), int(dt_str[5:7]), int(dt_str[8:10]))
        except:
            return ""
    else:
        dt = datetime.datetime.now()
    if freq == "daily":
        return dt.strftime("%Y-%m-%d")
    if freq == "weekly":
        # Week starts on due_day (0=Mon default)
        wd = due_day if due_day is not None else 0
        days_since = (dt.weekday() - wd) % 7
        start = dt - datetime.timedelta(days=days_since)
        return start.strftime("%Y-%m-%d")
    if freq == "monthly":
        # Period starts on due_day of month (default 1)
        dd = due_day if due_day and due_day >= 1 and due_day <= 28 else 1
        if dt.day >= dd:
            return dt.strftime("%Y-%m") + "-" + str(dd)
        else:
            # Before due day = still previous period
            prev = dt.month - 1
            yr = dt.year
            if prev < 1:
                prev = 12
                yr = yr - 1
            return str(yr) + "-" + str(prev).zfill(2) + "-" + str(dd)
    if freq == "annual":
        dm = due_month if due_month and due_month >= 1 and due_month <= 12 else 1
        dd = due_day if due_day and due_day >= 1 and due_day <= 28 else 1
        due_date = datetime.datetime(dt.year, dm, dd)
        if dt >= due_date:
            return str(dt.year) + "-" + str(dm).zfill(2) + "-" + str(dd)
        else:
            return str(dt.year - 1) + "-" + str(dm).zfill(2) + "-" + str(dd)
    return ""

# --- Default Groups (checks pulled from marketplace) -----------------
DEFAULT_GROUPS = [
    {"id":"data_health","name":"Data Health","desc":"Automated and manual data quality checks","icon":"&#128202;","color":"#0078d4","owner":"","best":True},
    {"id":"sunday_ops","name":"Sunday Operations","desc":"Weekly preparation for Sunday services","icon":"&#9961;","color":"#107c10","owner":"","best":True},
    {"id":"month_end","name":"Month-End Finance","desc":"Monthly financial close procedures","icon":"&#128176;","color":"#7a6400","owner":"","best":True},
    {"id":"new_member","name":"New Member Onboarding","desc":"Steps for welcoming and integrating new members","icon":"&#128100;","color":"#7c3aed","owner":"","best":True}
]

def parse_search_xml(xml):
    """Extract conditions from TouchPoint Search Builder XML into portable format."""
    import re
    conditions = []
    for m in re.finditer(r'<Condition[^>]*?Field="([^"]*)"[^>]*/>', xml):
        full = m.group(0)
        field = m.group(1)
        if field == 'Group':
            continue
        c = {"field": field}
        cm = re.search(r'Comparison="([^"]*)"', full)
        if cm:
            c["comparison"] = cm.group(1)
        vm = re.search(r'TextValue="([^"]*)"', full)
        if vm:
            c["value"] = vm.group(1)
        cv = re.search(r'CodeIdValue="([^"]*)"', full)
        if cv:
            c["value"] = cv.group(1)
        pm = re.search(r'Program="([^"]*)"', full)
        if pm and pm.group(1) != '0':
            c["program"] = pm.group(1)
        dm = re.search(r'Division="([^"]*)"', full)
        if dm and dm.group(1) != '0':
            c["division"] = dm.group(1)
        om = re.search(r'Organization="([^"]*)"', full)
        if om and om.group(1) != '0':
            c["organization"] = om.group(1)
        sm = re.search(r'Schedule="([^"]*)"', full)
        if sm:
            c["schedule"] = sm.group(1)
        conditions.append(c)
    return conditions

def conditions_to_code(conditions):
    """Convert parsed conditions to a portable Search Builder code string."""
    parts = []
    for c in conditions:
        field = c.get("field", "")
        comp = c.get("comparison", "Equal")
        val = c.get("value", "")
        if not field:
            continue
        # Skip church-specific org membership checks (have program IDs)
        if c.get("program") or c.get("organization"):
            continue
        op = "="
        if comp == "Greater":
            op = ">"
        elif comp == "Less":
            op = "<"
        elif comp == "GreaterEqual":
            op = ">="
        elif comp == "LessEqual":
            op = "<="
        elif comp == "NotEqual":
            op = "!="
        if val and ',' in val:
            bval = val.split(',')
            if len(bval) == 2 and bval[1] in ('True', 'False'):
                val = bval[0] + '[' + bval[1] + ']'
        parts.append(field + op + val)
    return " AND ".join(parts)

def lookup_search_by_name(search_name):
    """Look up a saved search by name and return its definition."""
    sql = "SELECT TOP 1 name, text, owner, ispublic FROM Query WHERE name = '" + search_name.replace("'", "''") + "'"
    try:
        for row in q.QuerySql(sql):
            xml = str(row.text) if row.text else ''
            conditions = parse_search_xml(xml)
            code = conditions_to_code(conditions)
            return {
                "found": True,
                "name": str(row.name),
                "owner": str(row.owner) if row.owner else '',
                "public": bool(row.ispublic) if row.ispublic else False,
                "conditions": conditions,
                "portable_code": code,
                "xml": xml
            }
    except:
        pass
    return {"found": False}

# --- Morning Batch: Email Reminders ----------------------------------
# Triggered either by TouchPoint's morning batch directly, or by a wrapper
# caller script that sets Data.run_batch = "true" and CallScript()s us.
def _is_batch_run():
    try:
        if model.FromMorningBatch:
            return True
    except:
        pass
    try:
        if hasattr(Data, 'run_batch') and str(Data.run_batch).lower() == 'true':
            return True
    except:
        pass
    return False

def send_reminder_emails():
    """Send today's reminder emails. Returns (sent, errors, message_str, log_lines).
    Used by both the morning batch trigger and the manual 'Send Now' action.
    Diagnostic messages are returned in log_lines (not printed) so AJAX callers
    don't pollute the JSON response body."""
    sent = 0
    errors = 0
    log_lines = []
    try:
        settings = lj(STK, "{}")
        default_email = settings.get("defaultEmail", "")
        default_pid = settings.get("defaultRecipientPid", 0)
        # Prefer PID-resolved email so a person's email change auto-propagates
        if default_pid:
            resolved = _resolve_pid_email(default_pid)
            if resolved:
                default_email = resolved
        cc_default = settings.get("ccDefault", False)
        if not default_email:
            return (0, 0, "No default recipient configured. Set it in Settings.", log_lines)
        dt = datetime.datetime.now()
        dow = dt.weekday()
        dom = dt.day
        moy = dt.month
        by_email = {}
        groups = lj(GK, "[]")
        results_cache = lj(RK, "{}")
        # Refresh automated/search check counts so the email reflects current state
        if isinstance(groups, list):
            t_refresh = now()
            for _gx in groups:
                try:
                    _checks_x = lj(gck(_gx["id"]), "[]")
                    if not isinstance(_checks_x, list):
                        continue
                    for _cx in _checks_x:
                        _tx = _cx.get("type", "manual")
                        if _tx == "auto" and _cx.get("sql"):
                            try:
                                _resx = run_sql_to_rows(_cx["sql"])
                                results_cache[_cx["id"]] = {"count": _resx["count"], "lastRun": t_refresh,
                                                            "detail": _resx["rows"], "columns": _resx["columns"],
                                                            "truncated": _resx["truncated"]}
                                if _resx["error"]:
                                    results_cache[_cx["id"]]["error"] = _resx["error"]
                            except Exception as _ea:
                                log_lines.append("Refresh failed for SQL check '" + str(_cx.get("name", "")) + "': " + str(_ea))
                        elif _tx == "search" and _cx.get("search"):
                            try:
                                _cntx = q.QueryCount(_cx["search"])
                                results_cache[_cx["id"]] = {"count": int(_cntx), "lastRun": t_refresh, "search": _cx["search"]}
                            except Exception as _es:
                                log_lines.append("Refresh failed for search check '" + str(_cx.get("name", "")) + "': " + str(_es))
                except:
                    pass
            sj(RK, results_cache)
        if isinstance(groups, list):
            log = lj(LK, "[]")
            if not isinstance(log, list):
                log = []
            ld = {}
            for e in log:
                ci = e.get("checkId", "")
                if ci and ci not in ld:
                    ld[ci] = e
            for g in groups:
                try:
                    checks = lj(gck(g["id"]), "[]")
                    if not isinstance(checks, list):
                        continue
                    for c in checks:
                        try:
                            freq = c.get("freq", "monthly")
                            dd = c.get("due_day")
                            dm = c.get("due_month")
                            is_due = False
                            if freq == "daily":
                                # Honor due_dows if set; default = every day for backward compat
                                ddws = c.get("due_dows", "")
                                if ddws:
                                    try:
                                        allowed = set([int(x.strip()) for x in str(ddws).split(',') if x.strip().isdigit()])
                                        is_due = (dow in allowed) if allowed else True
                                    except:
                                        is_due = True
                                else:
                                    is_due = True
                            elif freq == "weekly":
                                target = dd if dd is not None else 0
                                is_due = (dow == target)
                            elif freq == "monthly":
                                target = dd if dd and dd >= 1 else 1
                                is_due = (dom == target)
                            elif freq == "annual":
                                tm = dm if dm and dm >= 1 else 1
                                td = dd if dd and dd >= 1 else 1
                                is_due = (moy == tm and dom == td)
                            if not is_due:
                                continue
                            cid = c["id"]
                            if cid in ld:
                                entry_at = ld[cid].get("at", "")
                                if period_key(freq, entry_at, dd, dm) == period_key(freq, None, dd, dm):
                                    continue
                            recipients = []
                            # New: PeopleId-based recipients (resolved at send time)
                            ntpids_str = c.get("notifyPids", "")
                            if ntpids_str:
                                for pid_s in ntpids_str.split(','):
                                    pid_s = pid_s.strip()
                                    if not pid_s: continue
                                    try:
                                        em_resolved = _resolve_pid_email(int(pid_s))
                                        if em_resolved and em_resolved not in recipients:
                                            recipients.append(em_resolved)
                                    except:
                                        pass
                            # Legacy: comma-separated raw email strings
                            notify = c.get("notify", "")
                            if notify:
                                for em in notify.split(','):
                                    em = em.strip()
                                    if em and em not in recipients:
                                        recipients.append(em)
                            # CC default if any specific recipients exist
                            if recipients:
                                if cc_default and default_email and default_email not in recipients:
                                    recipients.append(default_email)
                            else:
                                # No specific recipients — fall back to default
                                if default_email:
                                    recipients.append(default_email)
                            # Look up cached count + last-run timestamp from results cache
                            _r = results_cache.get(c.get("id", ""), {})
                            _cnt = _r.get("count", -1)
                            _last = _r.get("lastRun", "")
                            for em in recipients:
                                if em not in by_email:
                                    by_email[em] = []
                                by_email[em].append({"group": g.get("name", ""), "icon": g.get("icon", ""),
                                                     "color": g.get("color", "#0078d4"),
                                                     "name": c.get("name", ""), "cat": c.get("cat", ""),
                                                     "freq": freq, "inst": c.get("inst", ""),
                                                     "type": c.get("type", "manual"),
                                                     "count": _cnt, "lastRun": _last})
                        except Exception as ce:
                            log_lines.append("Skipping check in group '" + str(g.get("name", "")) + "': " + str(ce))
                except Exception as ge:
                    log_lines.append("Skipping group '" + str(g.get("name", "")) + "': " + str(ge))
        if not by_email:
            return (0, 0, "Nothing due today. No emails sent.", log_lines)
        for email_addr in by_email:
            try:
                items = by_email[email_addr]
                if not items:
                    continue
                body = '<div style="font-family:Segoe UI,sans-serif;max-width:600px;margin:0 auto">'
                body = body + '<div style="background:linear-gradient(135deg,#0078d4,#005a9e);color:#fff;padding:20px 24px;border-radius:8px 8px 0 0">'
                body = body + '<h2 style="margin:0;font-size:20px">Operations Checklist Reminders</h2>'
                body = body + '<p style="margin:4px 0 0;opacity:.9;font-size:13px">' + str(len(items)) + ' items due today &mdash; ' + dt.strftime("%B %d, %Y") + '</p></div>'
                body = body + '<div style="background:#fff;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;padding:20px">'
                by_group = {}
                for item in items:
                    gn = item["group"]
                    if gn not in by_group:
                        by_group[gn] = []
                    by_group[gn].append(item)
                for gn in by_group:
                    gitems = by_group[gn]
                    body = body + '<h3 style="font-size:14px;color:#333;margin:16px 0 8px;border-bottom:1px solid #eee;padding-bottom:4px">' + safe(gn) + '</h3>'
                    for item in gitems:
                        type_badge = ''
                        count_badge = ''
                        # Count badge for automated/search checks based on cached results
                        item_type = item.get("type", "manual")
                        item_count = item.get("count", -1)
                        if item_type in ("auto", "search"):
                            if item_count > 0:
                                count_badge = ' <span style="background:#fde8e8;color:#d13438;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:700">' + str(item_count) + '</span>'
                            elif item_count == 0:
                                count_badge = ' <span style="background:#e8fde8;color:#107c10;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">&#10003; clean</span>'
                            else:
                                count_badge = ' <span style="background:#f0f0f0;color:#888;padding:2px 8px;border-radius:10px;font-size:11px">not run yet</span>'
                        if item_type == "auto":
                            type_badge = ' <span style="background:#e8fde8;color:#107c10;padding:1px 6px;border-radius:8px;font-size:10px;font-weight:600">SQL</span>'
                        elif item_type == "search":
                            type_badge = ' <span style="background:#f3e8fd;color:#7c3aed;padding:1px 6px;border-radius:8px;font-size:10px;font-weight:600">search</span>'
                        body = body + '<div style="padding:8px 12px;border-left:3px solid ' + safe(item["color"]) + ';margin-bottom:6px;background:#f8f9fa;border-radius:0 6px 6px 0">'
                        body = body + '<div style="font-size:13px;font-weight:600;color:#333">' + safe(item["name"]) + count_badge + type_badge + ' <span style="color:#888;font-weight:400;font-size:11px">' + safe(item["freq"]) + ' &middot; ' + safe(item["cat"]) + '</span></div>'
                        if item.get("lastRun"):
                            body = body + '<div style="font-size:10px;color:#999;margin-top:2px">as of ' + safe(item["lastRun"]) + '</div>'
                        if item["inst"]:
                            body = body + '<div style="font-size:12px;color:#666;margin-top:3px">' + safe(item["inst"]) + '</div>'
                        body = body + '</div>'
                body = body + '<div style="margin-top:20px;text-align:center">'
                body = body + '<a href="' + str(model.CmsHost) + '/PyScriptForm/' + get_script_name() + '" style="display:inline-block;padding:10px 24px;background:#0078d4;color:#fff;text-decoration:none;border-radius:6px;font-size:13px;font-weight:600">Open Checklists</a>'
                body = body + '</div></div></div>'
                pid = find_person_id_by_email(email_addr)
                if pid:
                    queued_by = pid
                    try:
                        upid = model.UserPeopleId
                        if upid:
                            queued_by = upid
                    except:
                        pass
                    subject = "Operations Checklist: " + str(len(items)) + " items due today"
                    model.Email(pid, queued_by, default_email, "Operations Checklists", subject, body)
                    sent = sent + 1
                else:
                    errors = errors + 1
                    log_lines.append("No TouchPoint person found for email '" + str(email_addr) + "' -- email NOT sent")
            except Exception as ee:
                errors = errors + 1
                log_lines.append("Failed to send to " + str(email_addr) + ": " + str(ee))
        msg = str(sent) + " reminder email" + ("s" if sent != 1 else "") + " sent"
        if errors:
            msg = msg + ", " + str(errors) + " failed"
        return (sent, errors, msg, log_lines)
    except Exception as me:
        return (sent, errors + 1, "Error: " + str(me), log_lines)

if _is_batch_run():
    # Always update lastBatchRun first — confirms the wrapper is calling us
    # regardless of whether sends are currently enabled, so the "batch not
    # active" banner clears as soon as the wrapper is wired up.
    try:
        settings_b0 = lj(STK, "{}")
        settings_b0["lastBatchRun"] = now()
        sj(STK, settings_b0)
    except:
        pass
    # Honor the in-app enable/disable toggle. Default = enabled.
    _settings_chk = lj(STK, "{}")
    if _settings_chk.get("batchEnabled", True):
        sent_b, err_b, msg_b, logs_b = send_reminder_emails()
        print "Ops Checklists: " + msg_b
        for ln_b in logs_b:
            print "Ops Checklists: " + ln_b
    else:
        print "Ops Checklists: Reminder sends are disabled in Settings; skipping."

# --- AJAX Handler Function -------------------------------------------
def handle_ajax():
    if model.HttpMethod != "post":
        return False
    if not hasattr(model.Data, 'action'):
        return False

    act = str(model.Data.action)

    # --- Capture & persist script name from any AJAX request ---------
    # JS sends window.location.pathname so we always have an authoritative name
    # regardless of model.URL behavior in IronPython.
    try:
        if hasattr(model.Data, 'script_name'):
            sn_in = str(model.Data.script_name).strip()
            if sn_in:
                _settings_sn = lj(STK, "{}")
                if _settings_sn.get("scriptName") != sn_in:
                    _settings_sn["scriptName"] = sn_in
                    sj(STK, _settings_sn)
    except:
        pass

    # --- Permission Gate ---------------------------------------------
    if act in _ADMIN_ACTIONS and not is_admin():
        print json.dumps({"success": False, "message": "Admin role required"})
        return True
    if act in _EDIT_ACTIONS and not can_edit():
        print json.dumps({"success": False, "message": "You don't have permission to edit checklists"})
        return True
    if act in _COMPLETE_ACTIONS and not can_complete():
        print json.dumps({"success": False, "message": "You don't have permission to update checks"})
        return True

    # --- Settings CRUD ------------------------------------------------
    if act == 'save_settings':
        de = str(model.Data.default_email) if hasattr(model.Data, 'default_email') else ''
        dpid_raw = str(model.Data.default_recipient_pid) if hasattr(model.Data, 'default_recipient_pid') else ''
        try:
            dpid = int(dpid_raw) if dpid_raw.strip() else 0
        except:
            dpid = 0
        cc = str(model.Data.cc_default) == 'true'
        er = str(model.Data.edit_roles) if hasattr(model.Data, 'edit_roles') else 'Admin'
        cr = str(model.Data.complete_roles) if hasattr(model.Data, 'complete_roles') else 'Admin'
        be = str(model.Data.batch_enabled) == 'true' if hasattr(model.Data, 'batch_enabled') else True
        settings = lj(STK, "{}")
        settings["defaultEmail"] = de
        settings["defaultRecipientPid"] = dpid
        settings["ccDefault"] = cc
        settings["editRoles"] = er.strip() or "Admin"
        settings["completeRoles"] = cr.strip() or "Admin"
        settings["batchEnabled"] = be
        sj(STK, settings)
        print json.dumps({"success": True})
        return True

    if act == 'get_settings':
        settings = lj(STK, "{}")
        # Provide sane defaults for the UI
        if "editRoles" not in settings:
            settings["editRoles"] = "Admin"
        if "completeRoles" not in settings:
            settings["completeRoles"] = "Admin"
        if "batchEnabled" not in settings:
            settings["batchEnabled"] = True
        print json.dumps({"success": True, "settings": settings, "isAdmin": is_admin()})
        return True

    if act == 'apply_update':
        # Fetch latest script from DisplayCache via workers.dev (bypasses CF
        # Bot Fight Mode that blocks server-side calls to the public domain)
        new_code = ''
        try:
            fetch_url = DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID
            new_code = str(model.RestGet(fetch_url, {}))
        except Exception as fe:
            print json.dumps({"success": False, "message": "Failed to fetch update: " + str(fe)})
            return True
        if not new_code or len(new_code) < 200:
            print json.dumps({"success": False, "message": "Invalid or empty script code received"})
            return True
        # Detect the actual script name TouchPoint installed this as
        target_name = get_script_name() or DC_SCRIPT_ID
        try:
            model.WriteContentPython(target_name, new_code)
            print json.dumps({"success": True, "message": "Updated " + target_name + ". Reload the page."})
        except Exception as we:
            print json.dumps({"success": False, "message": "Write failed: " + str(we)})
        return True

    if act == 'check_batch_install':
        # Report whether our managed block is in MorningBatch.
        try:
            existing_mb = model.PythonContent("MorningBatch") or ""
        except:
            existing_mb = ""
        installed = (_BATCH_MARKER_START in existing_mb)
        # Detect if any reference to our script exists outside our block,
        # so we can warn the admin instead of silently double-adding.
        sn_dbi = get_script_name()
        ref_outside = False
        try:
            stripped = existing_mb
            if installed:
                import re
                pat = re.escape(_BATCH_MARKER_START) + r".*?" + re.escape(_BATCH_MARKER_END)
                stripped = re.sub(pat, "", existing_mb, flags=re.DOTALL)
            if ('CallScript("' + sn_dbi + '")' in stripped) or ("CallScript('" + sn_dbi + "')" in stripped):
                ref_outside = True
        except:
            pass
        print json.dumps({"success": True, "installed": installed, "referencedOutsideBlock": ref_outside, "scriptName": sn_dbi})
        return True

    if act == 'install_batch':
        sn_ib = get_script_name()
        try:
            existing_mb = model.PythonContent("MorningBatch") or ""
        except Exception as ee_r:
            print json.dumps({"success": False, "message": "Could not read MorningBatch: " + str(ee_r)})
            return True
        if _BATCH_MARKER_START in existing_mb:
            print json.dumps({"success": True, "message": "Already installed in MorningBatch.", "alreadyInstalled": True})
            return True
        block = (
            _BATCH_MARKER_START + "\n"
            "try:\n"
            "    Data.run_batch = 'true'\n"
            "    model.CallScript('" + sn_ib + "')\n"
            "except Exception as _opschk_e:\n"
            "    print 'Ops Checklists batch error: ' + str(_opschk_e)\n"
            + _BATCH_MARKER_END + "\n"
        )
        new_content = (existing_mb.rstrip() + ("\n\n" if existing_mb.strip() else "") + block)
        try:
            model.WriteContentPython("MorningBatch", new_content)
        except Exception as ee_w:
            print json.dumps({"success": False, "message": "Could not write MorningBatch: " + str(ee_w)})
            return True
        print json.dumps({"success": True, "message": "Added to MorningBatch. Reminders will fire on the next morning batch run."})
        return True

    if act == 'uninstall_batch':
        try:
            existing_mb = model.PythonContent("MorningBatch") or ""
        except Exception as ee_r2:
            print json.dumps({"success": False, "message": "Could not read MorningBatch: " + str(ee_r2)})
            return True
        if _BATCH_MARKER_START not in existing_mb:
            print json.dumps({"success": True, "message": "Not currently installed in MorningBatch.", "notInstalled": True})
            return True
        import re
        pat = re.escape(_BATCH_MARKER_START) + r".*?" + re.escape(_BATCH_MARKER_END) + r"\n?"
        new_content = re.sub(pat, "", existing_mb, flags=re.DOTALL)
        # Tidy double-blank-lines left behind
        new_content = re.sub(r"\n{3,}", "\n\n", new_content).rstrip() + "\n"
        try:
            model.WriteContentPython("MorningBatch", new_content)
        except Exception as ee_w2:
            print json.dumps({"success": False, "message": "Could not write MorningBatch: " + str(ee_w2)})
            return True
        print json.dumps({"success": True, "message": "Removed from MorningBatch."})
        return True

    if act == 'send_reminders_now':
        sent_n, err_n, msg_n, logs_n = send_reminder_emails()
        # Record manual run timestamp separately so it doesn't suppress the
        # 'morning batch not active' banner (which only watches lastBatchRun)
        try:
            settings_mn = lj(STK, "{}")
            settings_mn["lastManualSend"] = now()
            settings_mn["lastManualSendBy"] = user()
            settings_mn["lastManualSendResult"] = msg_n
            sj(STK, settings_mn)
        except:
            pass
        print json.dumps({"success": True, "sent": sent_n, "errors": err_n, "message": msg_n, "log": logs_n})
        return True

    # --- Group CRUD --------------------------------------------------
    if act == 'save_group':
        gid = str(model.Data.group_id) if hasattr(model.Data, 'group_id') and str(model.Data.group_id) else ''
        nm = str(model.Data.group_name)
        desc = str(model.Data.group_desc) if hasattr(model.Data, 'group_desc') else ''
        icon = safe_icon(str(model.Data.group_icon)) if hasattr(model.Data, 'group_icon') else '&#128203;'
        clr = safe_color(str(model.Data.group_color)) if hasattr(model.Data, 'group_color') else '#0078d4'
        own = str(model.Data.group_owner) if hasattr(model.Data, 'group_owner') else ''

        groups = lj(GK, "[]")
        if not isinstance(groups, list):
            groups = []

        if gid:
            for g in groups:
                if g["id"] == gid:
                    g["name"] = nm
                    g["desc"] = desc
                    g["icon"] = icon
                    g["color"] = clr
                    g["owner"] = own
                    break
        else:
            gid = "g_" + uid()
            groups.append({"id": gid, "name": nm, "desc": desc, "icon": icon, "color": clr, "owner": own, "best": False})
            sj(gck(gid), [])

        sj(GK, groups)
        print json.dumps({"success": True, "id": gid})
        return True

    if act == 'delete_group':
        gid = str(model.Data.group_id)
        groups = lj(GK, "[]")
        if isinstance(groups, list):
            groups = [g for g in groups if g.get("id") != gid]
            sj(GK, groups)
        print json.dumps({"success": True})
        return True

    # --- Check CRUD (within group) -----------------------------------
    if act == 'save_check':
        gid = str(model.Data.group_id)
        cid = str(model.Data.check_id) if hasattr(model.Data, 'check_id') and str(model.Data.check_id) else ''
        nm = str(model.Data.check_name)
        ct = safe_cat(str(model.Data.check_cat)) if hasattr(model.Data, 'check_cat') else 'General'
        tp = safe_type(str(model.Data.check_type)) if hasattr(model.Data, 'check_type') else 'manual'
        sq = str(model.Data.check_sql) if hasattr(model.Data, 'check_sql') else ''
        sr = str(model.Data.check_search) if hasattr(model.Data, 'check_search') else ''
        fr = safe_freq(str(model.Data.check_freq)) if hasattr(model.Data, 'check_freq') else 'monthly'
        ins = str(model.Data.check_inst) if hasattr(model.Data, 'check_inst') else ''
        stp = str(model.Data.check_steps) if hasattr(model.Data, 'check_steps') else ''
        th = 0
        cr = 5
        dday = None
        dmonth = None
        try:
            th = int(model.Data.check_th)
        except:
            pass
        try:
            cr = int(model.Data.check_cr)
        except:
            pass
        try:
            dday = int(model.Data.check_due_day) if hasattr(model.Data, 'check_due_day') and str(model.Data.check_due_day) else None
        except:
            pass
        try:
            dmonth = int(model.Data.check_due_month) if hasattr(model.Data, 'check_due_month') and str(model.Data.check_due_month) else None
        except:
            pass
        ntfy = str(model.Data.check_notify) if hasattr(model.Data, 'check_notify') else ''
        # Person picker stores CSV of PeopleIds; we resolve to current emails at send-time
        ntfy_pids = str(model.Data.check_notify_pids) if hasattr(model.Data, 'check_notify_pids') else ''
        # Post-completion notify list — emailed when check is marked complete
        comp_notify_pids = str(model.Data.check_complete_notify_pids) if hasattr(model.Data, 'check_complete_notify_pids') else ''
        # Days-of-week for daily checks — CSV of integers 0=Mon..6=Sun.
        # New daily checks default to weekdays (Mon-Fri). Other freqs ignore.
        ddows = str(model.Data.check_due_dows) if hasattr(model.Data, 'check_due_dows') else ''

        checks = lj(gck(gid), "[]")
        if not isinstance(checks, list):
            checks = []

        sl = [s.strip() for s in stp.split('\n') if s.strip()] if stp else []

        # Default new daily checks to weekdays Mon-Fri unless explicitly set
        if fr == 'daily' and not ddows:
            ddows = '0,1,2,3,4'

        if cid:
            for c in checks:
                if c["id"] == cid:
                    c["name"] = nm
                    c["cat"] = ct
                    c["type"] = tp
                    c["sql"] = sq
                    c["search"] = sr
                    c["freq"] = fr
                    c["inst"] = ins
                    c["steps"] = sl
                    c["th"] = th
                    c["cr"] = cr
                    c["due_day"] = dday
                    c["due_month"] = dmonth
                    c["due_dows"] = ddows
                    c["notify"] = ntfy
                    c["notifyPids"] = ntfy_pids
                    c["completionNotifyPids"] = comp_notify_pids
                    break
        else:
            # Defensive: ensure new id doesn't collide with anything existing
            existing_ids = set(c.get("id", "") for c in checks if isinstance(c, dict))
            cid = "ck_" + uid()
            while cid in existing_ids:
                cid = "ck_" + uid()
            checks.append({"id": cid, "name": nm, "cat": ct, "type": tp, "sql": sq, "search": sr,
                           "freq": fr, "inst": ins, "steps": sl, "th": th, "cr": cr,
                           "due_day": dday, "due_month": dmonth, "due_dows": ddows,
                           "notify": ntfy, "notifyPids": ntfy_pids,
                           "completionNotifyPids": comp_notify_pids, "best": False})

        sj(gck(gid), checks)
        print json.dumps({"success": True, "id": cid})
        return True

    if act == 'dedupe_check_ids':
        # One-shot: walk all groups (or a single group) and reassign unique ids
        # to any duplicates. Run once if you experienced "delete-one-removes-many"
        # damage from the old second-precision uid().
        gid_scope = str(model.Data.group_id) if hasattr(model.Data, 'group_id') and str(model.Data.group_id) else ''
        groups_all = lj(GK, "[]")
        if not isinstance(groups_all, list):
            groups_all = []
        scanned = 0
        renamed = 0
        details = []
        for g in groups_all:
            if gid_scope and g.get("id") != gid_scope:
                continue
            checks = lj(gck(g["id"]), "[]")
            if not isinstance(checks, list):
                continue
            seen = set()
            dirty = False
            for c in checks:
                if not isinstance(c, dict):
                    continue
                scanned += 1
                cid = c.get("id", "")
                if cid in seen:
                    new_id = "ck_" + uid()
                    while new_id in seen:
                        new_id = "ck_" + uid()
                    details.append({"group": g.get("name", ""), "checkName": c.get("name", ""), "oldId": cid, "newId": new_id})
                    c["id"] = new_id
                    seen.add(new_id)
                    renamed += 1
                    dirty = True
                else:
                    seen.add(cid)
            if dirty:
                sj(gck(g["id"]), checks)
        print json.dumps({"success": True, "scanned": scanned, "renamed": renamed, "details": details})
        return True

    if act == 'delete_check':
        gid = str(model.Data.group_id)
        cid = str(model.Data.check_id)
        checks = lj(gck(gid), "[]")
        if isinstance(checks, list):
            # Only remove the FIRST match — protects against duplicate-ID
            # corruption from older second-precision uid() generation
            new_checks = []
            removed = False
            for c in checks:
                if not removed and c.get("id") == cid:
                    removed = True
                    continue
                new_checks.append(c)
            sj(gck(gid), new_checks)
        print json.dumps({"success": True})
        return True

    if act == 'update_all_marketplace':
        # Pull latest marketplace, then update every installed check whose
        # marketplace rev is newer than the local mp_rev.
        # Optional group_id scopes to a single group; otherwise updates across
        # all groups org-wide. If skip_edited is true, items the user has
        # locally modified are reported but not overwritten.
        gid_scope = str(model.Data.group_id) if hasattr(model.Data, 'group_id') and str(model.Data.group_id) else ''
        skip_edited = (str(model.Data.skip_edited).lower() == 'true') if hasattr(model.Data, 'skip_edited') else False
        mp_items, fetch_err = _get_marketplace_items_for_action()
        if fetch_err:
            print json.dumps({"success": False, "message": "Could not fetch marketplace: " + fetch_err})
            return True
        mp_by_name = {}
        for item in mp_items:
            mp_by_name[item.get("name", "")] = item
        groups_all = lj(GK, "[]")
        if not isinstance(groups_all, list):
            groups_all = []
        groups_to_scan = [g for g in groups_all if (not gid_scope or g.get("id") == gid_scope)]
        updated = 0
        scanned = 0
        skipped = 0
        details = []
        skipped_details = []
        # Track ids whose results need clearing (so stale count-only data
        # doesn't hide the ag-grid View button after SQL changes)
        cleared_cids = []
        for g in groups_to_scan:
            checks = lj(gck(g["id"]), "[]")
            if not isinstance(checks, list):
                continue
            dirty = False
            for c in checks:
                scanned += 1
                nm = c.get("name", "")
                tpl = mp_by_name.get(nm)
                if not tpl:
                    continue
                local_rev = c.get("mp_rev", 0) or 0
                mp_rev = tpl.get("rev", 1) or 1
                if mp_rev <= local_rev:
                    continue
                # Drift detection
                has_edits = _check_has_local_edits(c)
                if has_edits and skip_edited:
                    skipped += 1
                    skipped_details.append({"group": g.get("name", ""), "name": nm, "from": local_rev, "to": mp_rev})
                    continue
                c["name"] = tpl.get("name", c["name"])
                c["cat"] = tpl.get("cat", c.get("cat", "General"))
                c["type"] = tpl.get("type", c.get("type", "manual"))
                c["sql"] = tpl.get("sql", "")
                c["search"] = tpl.get("search", "")
                c["freq"] = tpl.get("freq", c.get("freq", "monthly"))
                c["inst"] = tpl.get("inst", "")
                c["steps"] = list(tpl.get("steps", []))
                c["th"] = tpl.get("th", 0)
                c["cr"] = tpl.get("cr", 5)
                c["mp_rev"] = mp_rev
                c["mp_snapshot"] = _make_mp_snapshot(tpl)
                dirty = True
                updated += 1
                cleared_cids.append(c["id"])
                details.append({"group": g.get("name", ""), "name": nm, "from": local_rev, "to": mp_rev, "hadEdits": has_edits})
            if dirty:
                sj(gck(g["id"]), checks)
        # Drop cached results for updated checks so the next page render shows
        # "not run yet" instead of stale count-only data hiding the View button.
        if cleared_cids:
            try:
                _rs_all = lj(RK, "{}")
                changed = False
                for _cc in cleared_cids:
                    if _cc in _rs_all:
                        del _rs_all[_cc]
                        changed = True
                if changed:
                    sj(RK, _rs_all)
            except:
                pass
        print json.dumps({
            "success": True,
            "updated": updated,
            "scanned": scanned,
            "skipped": skipped,
            "details": details,
            "skippedDetails": skipped_details
        })
        return True

    if act == 'update_from_marketplace':
        gid = str(model.Data.group_id)
        cid = str(model.Data.check_id)
        mp_name = str(model.Data.mp_name) if hasattr(model.Data, 'mp_name') else ''
        # Fetch marketplace to get latest version (browser-supplied if available)
        try:
            mp_items, fetch_err = _get_marketplace_items_for_action()
            if fetch_err:
                print json.dumps({"success": False, "message": "Could not fetch marketplace: " + fetch_err})
                return True
            if mp_items is not None:
                tpl = None
                for item in mp_items:
                    if item.get("name") == mp_name:
                        tpl = item
                        break
                if tpl:
                    checks = lj(gck(gid), "[]")
                    if isinstance(checks, list):
                        for c in checks:
                            if c["id"] == cid:
                                c["name"] = tpl.get("name", c["name"])
                                c["cat"] = tpl.get("cat", "General")
                                c["type"] = tpl.get("type", "manual")
                                c["sql"] = tpl.get("sql", "")
                                c["search"] = tpl.get("search", "")
                                c["freq"] = tpl.get("freq", "monthly")
                                c["inst"] = tpl.get("inst", "")
                                c["steps"] = list(tpl.get("steps", []))
                                c["th"] = tpl.get("th", 0)
                                c["cr"] = tpl.get("cr", 5)
                                c["mp_rev"] = tpl.get("rev", 1)
                                c["mp_snapshot"] = _make_mp_snapshot(tpl)
                                break
                        sj(gck(gid), checks)
                        # Stale cached results would hide the new ag-grid View
                        # button (old count-only results have no detail rows).
                        try:
                            _rs = lj(RK, "{}")
                            if cid in _rs:
                                del _rs[cid]
                                sj(RK, _rs)
                        except:
                            pass
                    print json.dumps({"success": True})
                    return True
        except:
            pass
        print json.dumps({"success": False, "message": "Could not fetch update"})
        return True

    # --- Library sync (enable/disable catalog checks) ------------------
    if act == 'sync_library':
        gid = str(model.Data.group_id)
        enabled_raw = str(model.Data.enabled_indexes) if hasattr(model.Data, 'enabled_indexes') else ''
        enabled_idxs = set()
        for x in enabled_raw.split(','):
            x = x.strip()
            if x:
                try:
                    enabled_idxs.add(int(x))
                except:
                    pass
        # Fetch marketplace catalog (browser-supplied if available, else CF)
        mp_items, _fe = _get_marketplace_items_for_action()
        if not mp_items and _fe:
            print json.dumps({"success": False, "message": "Could not fetch marketplace: " + _fe})
            return True

        checks = lj(gck(gid), "[]")
        if not isinstance(checks, list):
            checks = []
        existing_names = {}
        for c in checks:
            existing_names[c.get("name", "")] = c
        # Add newly enabled items (indexes match JS CATALOG from marketplace)
        added = 0
        enabled_names = set()
        for idx in enabled_idxs:
            if idx >= 0 and idx < len(mp_items):
                tpl = mp_items[idx]
                enabled_names.add(tpl.get("name", ""))
                if tpl.get("name", "") not in existing_names:
                    cid = "ck_" + uid() + "_" + str(added)
                    # Use portable code instead of church-specific search name
                    search_val = tpl.get("search", "")
                    if tpl.get("portableCode"):
                        search_val = tpl["portableCode"]
                    # Build instructions with condition details if available
                    inst = tpl.get("inst", "")
                    if tpl.get("searchConditions"):
                        cond_lines = []
                        for sc in tpl["searchConditions"]:
                            line = sc.get("field", "") + " " + sc.get("comparison", "") + " " + sc.get("value", "")
                            if sc.get("program"):
                                line = line + " (Program: " + sc["program"] + ")"
                            if sc.get("division"):
                                line = line + " (Division: " + sc["division"] + ")"
                            cond_lines.append(line)
                        if cond_lines:
                            inst = inst + "\n\nOriginal conditions: " + "; ".join(cond_lines)
                    checks.append({"id": cid, "name": tpl.get("name", ""), "cat": tpl.get("cat", "General"),
                                   "type": tpl.get("type", "manual"), "sql": tpl.get("sql", ""),
                                   "search": search_val, "freq": tpl.get("freq", "monthly"),
                                   "inst": inst, "steps": list(tpl.get("steps", [])),
                                   "th": tpl.get("th", 0), "cr": tpl.get("cr", 5), "best": False,
                                   "src": "catalog", "mp_rev": tpl.get("rev", 1),
                                   "mp_snapshot": _make_mp_snapshot(tpl)})
                    added += 1
        # Remove disabled checks whose name matches any marketplace item
        catalog_names = set()
        for item in mp_items:
            catalog_names.add(item.get("name", ""))
        removed = 0
        new_checks = []
        for c in checks:
            cn = c.get("name", "")
            if cn in catalog_names and cn not in enabled_names:
                removed += 1
            else:
                new_checks.append(c)
        sj(gck(gid), new_checks)
        print json.dumps({"success": True, "added": added, "removed": removed})
        return True

    # --- Run Checks --------------------------------------------------
    if act == 'run_checks':
        gid = str(model.Data.group_id)
        checks = lj(gck(gid), "[]")
        if not isinstance(checks, list):
            checks = []
        rs = lj(RK, "{}")
        t = now()
        for c in checks:
            ctype = c.get("type", "manual")
            if ctype == "auto" and c.get("sql"):
                res = run_sql_to_rows(c["sql"])
                rs[c["id"]] = {"count": res["count"], "lastRun": t,
                               "detail": res["rows"], "columns": res["columns"],
                               "truncated": res["truncated"]}
                if res["error"]:
                    rs[c["id"]]["error"] = res["error"]
            elif ctype == "search" and c.get("search"):
                sres = run_search_to_rows(c["search"])
                rs[c["id"]] = {"count": sres["count"], "lastRun": t, "search": c["search"],
                               "detail": sres["rows"], "columns": sres["columns"],
                               "truncated": sres["truncated"]}
                if sres["error"]:
                    rs[c["id"]]["error"] = sres["error"]
        sj(RK, rs)
        print json.dumps({"success": True, "timestamp": t})
        return True

    if act == 'run_single':
        cid = str(model.Data.check_id)
        gid = str(model.Data.group_id)
        checks = lj(gck(gid), "[]")
        if not isinstance(checks, list):
            checks = []
        rs = lj(RK, "{}")
        t = now()
        for c in checks:
            if c["id"] == cid:
                ctype = c.get("type", "manual")
                if ctype == "auto" and c.get("sql"):
                    res = run_sql_to_rows(c["sql"])
                    rs[cid] = {"count": res["count"], "lastRun": t,
                               "detail": res["rows"], "columns": res["columns"],
                               "truncated": res["truncated"]}
                    if res["error"]:
                        rs[cid]["error"] = res["error"]
                elif ctype == "search" and c.get("search"):
                    sres = run_search_to_rows(c["search"])
                    rs[cid] = {"count": sres["count"], "lastRun": t, "search": c["search"],
                               "detail": sres["rows"], "columns": sres["columns"],
                               "truncated": sres["truncated"]}
                    if sres["error"]:
                        rs[cid]["error"] = sres["error"]
                break
        sj(RK, rs)
        print json.dumps({"success": True, "count": rs.get(cid, {}).get("count", -1), "lastRun": t})
        return True

    if act == 'bulk_tag':
        pids_raw = str(model.Data.people_ids) if hasattr(model.Data, 'people_ids') else ''
        tag_name = str(model.Data.tag_name) if hasattr(model.Data, 'tag_name') else ''
        tag_name = tag_name.strip()
        if not pids_raw or not tag_name:
            print json.dumps({"success": False, "message": "Missing PeopleIds or tag name"})
            return True
        try:
            pids = []
            for p in pids_raw.split(','):
                p = p.strip()
                if p:
                    try:
                        pids.append(int(p))
                    except:
                        pass
            if not pids:
                print json.dumps({"success": False, "message": "No valid PeopleIds"})
                return True
            pid_csv = ','.join(str(p) for p in pids)
            query_str = "peopleids='" + pid_csv + "'"
            current_user = None
            try:
                current_user = model.UserPeopleId
            except:
                pass
            model.AddTag(query_str, tag_name, current_user, False)
            print json.dumps({"success": True, "count": len(pids),
                              "message": "Tagged " + str(len(pids)) + " people with \"" + tag_name + "\""})
        except Exception as e:
            print json.dumps({"success": False, "message": "Tag failed: " + str(e)})
        return True

    if act == 'search_people':
        qstr = str(model.Data.q).strip() if hasattr(model.Data, 'q') else ''
        if len(qstr) < 2:
            print json.dumps({"success": True, "results": []})
            return True
        # Build WHERE clause:
        #   Single word: match Name (First Last), Name2 (Last, First), FirstName,
        #     LastName, NickName, EmailAddress, EmailAddress2
        #   Multi-word: each word must match a name field — handles both
        #     "Ben Swaby" (First Last) and "Swaby Ben" (Last First)
        parts = [p for p in qstr.split() if p]
        if len(parts) > 1:
            clauses = []
            for part in parts:
                sp = part.replace("'", "''")
                clauses.append(
                    "(p.FirstName LIKE '%" + sp + "%' "
                    "OR p.LastName LIKE '%" + sp + "%' "
                    "OR p.NickName LIKE '%" + sp + "%' "
                    "OR p.PreferredName LIKE '%" + sp + "%')"
                )
            where_clause = " AND ".join(clauses)
        else:
            safe_q = qstr.replace("'", "''")
            where_clause = (
                "(p.Name LIKE '%" + safe_q + "%' "
                "OR p.Name2 LIKE '%" + safe_q + "%' "
                "OR p.FirstName LIKE '%" + safe_q + "%' "
                "OR p.LastName LIKE '%" + safe_q + "%' "
                "OR p.NickName LIKE '%" + safe_q + "%' "
                "OR p.EmailAddress LIKE '%" + safe_q + "%' "
                "OR p.EmailAddress2 LIKE '%" + safe_q + "%')"
            )
        sql = (
            "SELECT TOP 25 p.PeopleId, p.Name2 as Name, p.EmailAddress, p.EmailAddress2, "
            "ms.Description as MemberStatus "
            "FROM People p "
            "LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id "
            "WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0 "
            "AND " + where_clause + " "
            "ORDER BY p.Name2"
        )
        try:
            results = []
            for row in q.QuerySql(sql):
                em = _safe_str(row.EmailAddress) or _safe_str(row.EmailAddress2)
                results.append({
                    "PeopleId": int(row.PeopleId),
                    "Name": _safe_str(row.Name),
                    "Email": em,
                    "MemberStatus": _safe_str(row.MemberStatus)
                })
            print json.dumps({"success": True, "results": results})
        except Exception as e:
            print json.dumps({"success": False, "message": str(e)})
        return True

    if act == 'search_saved_searches':
        qstr = str(model.Data.q).strip() if hasattr(model.Data, 'q') else ''
        if len(qstr) < 2:
            print json.dumps({"success": True, "results": []})
            return True
        safe_q = qstr.replace("'", "''")
        sql = (
            "SELECT TOP 25 Name, Owner, IsPublic "
            "FROM Query "
            "WHERE Name LIKE '%" + safe_q + "%' "
            "ORDER BY CASE WHEN Name LIKE '" + safe_q + "%' THEN 0 ELSE 1 END, Name"
        )
        try:
            results = []
            for row in q.QuerySql(sql):
                ipub = False
                try:
                    ipub = bool(row.IsPublic)
                except:
                    pass
                results.append({
                    "Name": _safe_str(row.Name),
                    "Owner": _safe_str(row.Owner) if hasattr(row, 'Owner') else '',
                    "Public": ipub
                })
            print json.dumps({"success": True, "results": results})
        except Exception as e:
            print json.dumps({"success": False, "message": str(e)})
        return True

    if act == 'resolve_people':
        # Given a CSV of PeopleIds, return Name + Email for each (for picker
        # chip display when re-opening a saved check/settings).
        pids_raw = str(model.Data.pids) if hasattr(model.Data, 'pids') else ''
        pids = []
        for p in pids_raw.split(','):
            p = p.strip()
            if p:
                try: pids.append(int(p))
                except: pass
        if not pids:
            print json.dumps({"success": True, "results": []})
            return True
        sql = (
            "SELECT p.PeopleId, p.Name2 as Name, p.EmailAddress, p.EmailAddress2 "
            "FROM People p "
            "WHERE p.PeopleId IN (" + ','.join(str(x) for x in pids) + ")"
        )
        try:
            by_pid = {}
            for row in q.QuerySql(sql):
                em = _safe_str(row.EmailAddress) or _safe_str(row.EmailAddress2)
                by_pid[int(row.PeopleId)] = {
                    "PeopleId": int(row.PeopleId),
                    "Name": _safe_str(row.Name),
                    "Email": em
                }
            # Return in the same order the caller asked
            results = [by_pid[p] for p in pids if p in by_pid]
            print json.dumps({"success": True, "results": results})
        except Exception as e:
            print json.dumps({"success": False, "message": str(e)})
        return True

    if act == 'get_check_detail':
        cid = str(model.Data.check_id)
        rs = lj(RK, "{}")
        det = rs.get(cid, {})
        print json.dumps({
            "success": True,
            "rows": det.get("detail", []),
            "columns": det.get("columns", []),
            "count": det.get("count", -1),
            "truncated": det.get("truncated", False),
            "lastRun": det.get("lastRun", ""),
            "error": det.get("error", "")
        })
        return True

    # --- Submit to marketplace ----------------------------------------
    if act == 'submit_to_marketplace':
        nm = str(model.Data.sub_name) if hasattr(model.Data, 'sub_name') else ''
        ct = str(model.Data.sub_cat) if hasattr(model.Data, 'sub_cat') else 'General'
        fr = str(model.Data.sub_freq) if hasattr(model.Data, 'sub_freq') else 'monthly'
        tp = str(model.Data.sub_type) if hasattr(model.Data, 'sub_type') else 'manual'
        sq = str(model.Data.sub_sql) if hasattr(model.Data, 'sub_sql') else ''
        sr = str(model.Data.sub_search) if hasattr(model.Data, 'sub_search') else ''
        ins = str(model.Data.sub_inst) if hasattr(model.Data, 'sub_inst') else ''
        stp = str(model.Data.sub_steps) if hasattr(model.Data, 'sub_steps') else ''
        sl = [s.strip() for s in stp.split('\n') if s.strip()] if stp else []
        th = 0
        cr = 5
        try:
            th = int(model.Data.sub_th)
        except:
            pass
        try:
            cr = int(model.Data.sub_cr)
        except:
            pass
        church = ''
        try:
            church = str(model.Setting("NameOfChurch"))
        except:
            pass
        # If search type, auto-extract conditions from saved search
        search_conditions = []
        portable_code = ''
        if tp == 'search' and sr:
            defn = lookup_search_by_name(sr)
            if defn.get("found"):
                search_conditions = defn.get("conditions", [])
                portable_code = defn.get("portable_code", "")
        payload = json.dumps({
            "name": nm, "cat": ct, "freq": fr, "checkType": tp,
            "sql": sq, "search": sr, "inst": ins, "steps": sl,
            "th": th, "cr": cr, "submittedBy": user(), "church": church,
            "searchConditions": search_conditions,
            "portableCode": portable_code
        })
        try:
            resp = model.RestPost("https://scripts.displaycache.com/api/ops-catalog/submit",
                                  {"Content-Type": "application/json"}, payload)
            result = json.loads(resp) if resp else {"success": False}
            print json.dumps(result)
        except Exception as e:
            print json.dumps({"success": False, "message": str(e)})
        return True

    # --- Saved search lookup ------------------------------------------
    if act == 'lookup_search':
        name = str(model.Data.search_name) if hasattr(model.Data, 'search_name') else ''
        if not name:
            print json.dumps({"success": False, "message": "Search name required"})
            return True
        defn = lookup_search_by_name(name)
        if defn.get("found"):
            defn["success"] = True
            print json.dumps(defn)
        else:
            print json.dumps({"success": False, "message": "Search '" + name + "' not found"})
        return True

    # --- Steps & Completion ------------------------------------------
    if act == 'toggle_step':
        cid = str(model.Data.check_id)
        si = str(model.Data.step_idx)
        chk = str(model.Data.checked) == 'true'
        nt = str(model.Data.note) if hasattr(model.Data, 'note') else ''
        sd = lj(SK, "{}")
        if cid not in sd:
            sd[cid] = {}
        sd[cid][si] = {"done": chk, "note": nt, "by": user(), "at": now()}
        sj(SK, sd)
        print json.dumps({"success": True})
        return True

    if act == 'save_step_note':
        cid = str(model.Data.check_id)
        si = str(model.Data.step_idx)
        nt = str(model.Data.note)
        sd = lj(SK, "{}")
        if cid not in sd:
            sd[cid] = {}
        ex = sd[cid].get(si, {"done": False})
        ex["note"] = nt
        ex["by"] = user()
        ex["at"] = now()
        sd[cid][si] = ex
        sj(SK, sd)
        print json.dumps({"success": True})
        return True

    if act == 'complete_check':
        cid = str(model.Data.check_id)
        gid = str(model.Data.group_id) if hasattr(model.Data, 'group_id') else ''
        nts = str(model.Data.notes) if hasattr(model.Data, 'notes') else ''
        completed_at = now()
        completed_by = user()
        lg = lj(LK, "[]")
        if not isinstance(lg, list):
            lg = []
        lg.insert(0, {"checkId": cid, "groupId": gid, "by": completed_by, "at": completed_at, "notes": nts})
        lg = lg[:500]
        sj(LK, lg)
        sd = lj(SK, "{}")
        if cid in sd:
            del sd[cid]
        sj(SK, sd)
        # Send post-completion notifications (silent on failure so it never
        # blocks the actual completion response)
        notif_sent = 0
        try:
            check_obj = None
            group_obj = None
            if gid:
                checks = lj(gck(gid), "[]")
                if isinstance(checks, list):
                    for c in checks:
                        if c.get("id") == cid:
                            check_obj = c
                            break
                groups_all = lj(GK, "[]")
                if isinstance(groups_all, list):
                    for g in groups_all:
                        if g.get("id") == gid:
                            group_obj = g
                            break
            cnpids = check_obj.get("completionNotifyPids", "") if check_obj else ""
            if cnpids:
                notif_sent = _send_completion_notice(check_obj, group_obj, completed_by, completed_at, nts, cnpids)
        except:
            pass
        print json.dumps({"success": True, "at": completed_at, "by": completed_by, "notified": notif_sent})
        return True

    # --- First-run setup -----------------------------------------------
    if act == 'setup':
        selected = str(model.Data.selected_groups) if hasattr(model.Data, 'selected_groups') else ''
        group_ids = [x.strip() for x in selected.split(',') if x.strip()]
        # Fetch default checks from marketplace (browser-supplied if available)
        mp_defaults = {}
        try:
            _setup_items, _setup_err = _get_marketplace_items_for_action()
            if not _setup_err:
                for item in _setup_items:
                    if item.get("default") and item.get("defaultGroup"):
                        dg = item["defaultGroup"]
                        if dg not in mp_defaults:
                            mp_defaults[dg] = []
                        cid = "ck_" + uid() + "_" + str(len(mp_defaults[dg]))
                        mp_defaults[dg].append({
                            "id": cid, "name": item.get("name", ""), "cat": item.get("cat", "General"),
                            "type": item.get("type", "manual"), "sql": item.get("sql", ""),
                            "search": item.get("search", ""), "freq": item.get("freq", "monthly"),
                            "inst": item.get("inst", ""), "steps": list(item.get("steps", [])),
                            "th": item.get("th", 0), "cr": item.get("cr", 5), "best": True,
                            "src": "catalog", "mp_rev": item.get("rev", 1),
                            "mp_snapshot": _make_mp_snapshot(item)
                        })
        except:
            pass
        groups = []
        for g in DEFAULT_GROUPS:
            if g["id"] in group_ids:
                groups.append(g)
                dc = mp_defaults.get(g["id"], [])
                sj(gck(g["id"]), dc)
        if not groups:
            groups = []
        sj(GK, groups)
        sj("TPxi_OpsChecklist_SetupDone", {"done": True})
        print json.dumps({"success": True, "count": len(groups)})
        return True

    if act == 'reset_group':
        gid = str(model.Data.group_id)
        checks = lj(gck(gid), "[]")
        if not isinstance(checks, list):
            checks = []
        sd = lj(SK, "{}")
        for c in checks:
            if c["id"] in sd:
                del sd[c["id"]]
        sj(SK, sd)
        print json.dumps({"success": True})
        return True

    return False

# --- Run AJAX handler ------------------------------------------------
_ajax = handle_ajax()

# --- Render Page (module level) --------------------------------------
if not _ajax:
    groups = lj(GK, "[]")
    if not isinstance(groups, list):
        groups = []
    _setup = False

    # Persist the detected script name so the morning batch (which has no URL)
    # can render the correct link in reminder emails.
    try:
        _detected_name = get_script_name()
        _settings_sn = lj(STK, "{}")
        if _settings_sn.get("scriptName") != _detected_name:
            _settings_sn["scriptName"] = _detected_name
            sj(STK, _settings_sn)
    except:
        pass

    # --- First-run setup wizard --------------------------------------
    setup_done = lj("TPxi_OpsChecklist_SetupDone", "{}")
    if len(groups) == 0 and not setup_done.get("done"):
        setup_html = '''
<style>
*{box-sizing:border-box}
.setup{max-width:700px;margin:40px auto;font-family:Segoe UI,-apple-system,sans-serif;padding:0 20px}
.setup-head{text-align:center;margin-bottom:32px}
.setup-head h2{font-size:26px;color:#333;margin:0 0 8px}
.setup-head p{font-size:14px;color:#888;line-height:1.5;max-width:500px;margin:0 auto}
.setup-step{display:flex;align-items:center;gap:12px;margin-bottom:24px;padding-bottom:24px;border-bottom:1px solid #eee}
.setup-num{width:32px;height:32px;border-radius:50%;background:#0078d4;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;flex-shrink:0}
.setup-txt{font-size:13px;color:#555;line-height:1.4}
.setup-txt strong{color:#333}
.sg{border:1px solid #e0e0e0;border-radius:12px;padding:16px;margin-bottom:12px;cursor:pointer;transition:all .2s;display:flex;gap:14px;align-items:flex-start}
.sg:hover{box-shadow:0 2px 12px rgba(0,0,0,.08)}
.sg.on{border-color:#0078d4;background:#f0f7ff}
.sg-icon{width:48px;height:48px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:24px;flex-shrink:0}
.sg-info{flex:1}
.sg-nm{font-size:15px;font-weight:700;color:#333;margin-bottom:2px}
.sg-desc{font-size:12px;color:#888;line-height:1.4}
.sg-checks{font-size:11px;color:#0078d4;margin-top:4px;font-weight:600}
.sg-toggle{width:20px;height:20px;border:2px solid #ccc;border-radius:4px;flex-shrink:0;display:flex;align-items:center;justify-content:center;margin-top:2px;transition:all .2s}
.sg.on .sg-toggle{background:#0078d4;border-color:#0078d4;color:#fff}
.setup-actions{text-align:center;margin-top:24px}
.setup-btn{padding:12px 32px;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;transition:all .15s}
.setup-primary{background:#0078d4;color:#fff}
.setup-primary:hover{background:#005a9e}
.setup-skip{background:none;border:none;color:#888;font-size:13px;cursor:pointer;margin-top:12px;display:block;margin-left:auto;margin-right:auto}
.setup-skip:hover{color:#555}
</style>

<div class="setup">
    <div class="setup-head">
        <div style="font-size:48px;margin-bottom:12px">&#9889;</div>
        <h2>Welcome to Operations Checklists</h2>
        <p>Keep your church data healthy and operations running smoothly with automated checks, team checklists, and best practices.</p>
    </div>

    <div class="setup-step">
        <div class="setup-num">1</div>
        <div class="setup-txt"><strong>Choose your checklist groups</strong> &mdash; Select the areas you want to manage. You can always add more later.</div>
    </div>
    <div class="setup-step">
        <div class="setup-num">2</div>
        <div class="setup-txt"><strong>Each group comes with recommended checks</strong> &mdash; Automated SQL queries, saved searches, and manual steps curated by the TouchPoint community.</div>
    </div>
    <div class="setup-step">
        <div class="setup-num">3</div>
        <div class="setup-txt"><strong>Customize from the library</strong> &mdash; Add, remove, or edit checks. Browse the marketplace for more. Share your own with other churches.</div>
    </div>

    <div id="setup_groups"></div>

    <div class="setup-actions">
        <button class="setup-btn setup-primary" onclick="finishSetup()">Get Started</button>
        <button class="setup-skip" onclick="skipSetup()">Start empty &mdash; I'll build my own</button>
    </div>
</div>
'''
        model.Form = setup_html

        def js_safe_s(obj):
            s = json.dumps(obj)
            s = s.replace('\\', '\\\\')
            s = s.replace("'", "\\'")
            s = s.replace('</', '<\\/')
            return s

        setup_js = """
var DG = JSON.parse('""" + js_safe_s(DEFAULT_GROUPS) + """');
var DC = {};
var SETUP_CATALOG = [];  // Cached so finishSetup can POST it (avoids server-side CF block)
var selected = {};
DG.forEach(function(g) { selected[g.id] = true; });

// Fetch default counts from marketplace (and cache full catalog for setup POST)
$.get('https://scripts.displaycache.com/api/ops-catalog', function(r) {
    try {
        var data = typeof r === 'string' ? JSON.parse(r) : r;
        SETUP_CATALOG = data.items || [];
        SETUP_CATALOG.forEach(function(item) {
            if (item.default && item.defaultGroup) {
                DC[item.defaultGroup] = (DC[item.defaultGroup] || 0) + 1;
            }
        });
    } catch(e) {}
    renderSetupGroups();
});

function renderSetupGroups() {
    var h = '';
    DG.forEach(function(g) {
        var on = selected[g.id] ? ' on' : '';
        var cnt = DC[g.id] || 0;
        h += '<div class="sg' + on + '" onclick="toggleSetupGroup(\\'' + g.id + '\\')">';
        h += '<div class="sg-icon" style="background:' + g.color + '20;color:' + g.color + '">' + g.icon + '</div>';
        h += '<div class="sg-info"><div class="sg-nm">' + g.name + '</div>';
        h += '<div class="sg-desc">' + g.desc + '</div>';
        h += '<div class="sg-checks">' + (cnt > 0 ? cnt + ' recommended checks included' : 'Loading...') + '</div></div>';
        h += '<div class="sg-toggle">' + (selected[g.id] ? '&#10003;' : '') + '</div>';
        h += '</div>';
    });
    document.getElementById('setup_groups').innerHTML = h;
}

function toggleSetupGroup(id) {
    selected[id] = !selected[id];
    renderSetupGroups();
}

function finishSetup() {
    var ids = [];
    for (var k in selected) { if (selected[k]) ids.push(k); }
    if (ids.length === 0) { alert('Select at least one group, or click "Start empty" below.'); return; }
    var btn = event.target;
    btn.disabled = true; btn.textContent = 'Setting up...';
    $.post(window.location.pathname, {
        action: 'setup',
        selected_groups: ids.join(','),
        catalog_json: JSON.stringify({items: SETUP_CATALOG})
    }, function(r) {
        try { var d = JSON.parse(r); if (d.success) location.reload(); } catch(e) {}
        btn.disabled = false; btn.textContent = 'Get Started';
    });
}

function skipSetup() {
    $.post(window.location.pathname, {
        action: 'setup',
        selected_groups: '',
        catalog_json: JSON.stringify({items: SETUP_CATALOG})
    }, function(r) {
        location.reload();
    });
}

renderSetupGroups();
"""
        model.Script = setup_js
        _setup = True

    if not _setup:
        # --- Normal render (existing groups) ---
        results = lj(RK, "{}")
        log = lj(LK, "[]")
        sd = lj(SK, "{}")
        if not isinstance(log, list):
            log = []

        # Most recent completion per check (all time)
        ld = {}
        for e in log:
            ci = e.get("checkId", "")
            if ci and ci not in ld:
                ld[ci] = e

        # Load all checks, backfill missing 'cat' (v1.1 migration)
        all_checks = {}
        check_freq = {}
        check_due = {}
        for g in groups:
            checks = lj(gck(g["id"]), "[]")
            if not isinstance(checks, list):
                checks = []
            dirty = False
            for c in checks:
                check_freq[c["id"]] = c.get("freq", "monthly")
                check_due[c["id"]] = {"day": c.get("due_day"), "month": c.get("due_month")}
                # Backfill missing category from catalog/defaults
                if not c.get("cat"):
                    c["cat"] = "General"
                    dirty = True
            if dirty:
                sj(gck(g["id"]), checks)
            all_checks[g["id"]] = checks

        # Current-period completions only
        ld_current = {}
        for cid in ld:
            freq = check_freq.get(cid, "monthly")
            dd = check_due.get(cid, {})
            entry_at = ld[cid].get("at", "")
            if period_key(freq, entry_at, dd.get("day"), dd.get("month")) == period_key(freq, None, dd.get("day"), dd.get("month")):
                ld_current[cid] = ld[cid]

        # Clear stale step progress (steps from a previous period)
        sd_dirty = False
        for cid in list(sd.keys()):
            freq = check_freq.get(cid, "monthly")
            dd = check_due.get(cid, {})
            if cid in ld and cid not in ld_current:
                del sd[cid]
                sd_dirty = True
            elif cid in sd and sd[cid]:
                first_step = sd[cid].get("0", {})
                step_at = first_step.get("at", "")
                if step_at and period_key(freq, step_at, dd.get("day"), dd.get("month")) != period_key(freq, None, dd.get("day"), dd.get("month")):
                    del sd[cid]
                    sd_dirty = True
        if sd_dirty:
            sj(SK, sd)

        # Build group stats using period-aware completion
        gstats = {}
        for g in groups:
            checks = all_checks[g["id"]]
            total = len(checks)
            done = 0
            issues = 0
            ok_count = 0
            warn_count = 0
            crit_count = 0
            for c in checks:
                r = results.get(c["id"], {})
                cn = r.get("count", -1)
                if c["type"] in ("auto", "search") and cn >= 0:
                    if cn == 0:
                        ok_count += 1
                    elif cn >= c.get("cr", 5):
                        crit_count += 1
                        issues += cn
                    elif cn > c.get("th", 0):
                        warn_count += 1
                        issues += cn
                if c["id"] in ld_current:
                    done += 1
                elif c["type"] in ("auto", "search") and cn == 0:
                    done += 1
            gstats[g["id"]] = {"total": total, "done": done, "issues": issues,
                                "ok": ok_count, "warn": warn_count, "crit": crit_count}

        def js_safe(obj):
            s = json.dumps(obj)
            s = s.replace('\\', '\\\\')
            s = s.replace("'", "\\'")
            s = s.replace('</', '<\\/')
            return s

        gj = js_safe(groups)
        acj = js_safe(all_checks)
        rj = js_safe(results)
        ldj = js_safe(ld)
        lcj = js_safe(ld_current)
        sdj = js_safe(sd)
        gsj = js_safe(gstats)
        # Full log (capped to 500 server-side) for calendar past-view
        logj = js_safe(log if isinstance(log, list) else [])

        # User capabilities (computed server-side, enforced in handle_ajax)
        _can_edit = can_edit()
        _can_complete = can_complete()
        _is_admin = is_admin()

        # Morning batch wiring check.
        # Logic:
        #   - Installed in MorningBatch (managed block OR any reference)
        #     - AND has run recently (<36h) → all good, no banner
        #     - AND has run but stale → "stale" warning
        #     - AND has never run → "installed, waiting for next morning batch" info banner
        #   - Not installed AND has run (manual wrapper) → no banner
        #   - Not installed AND never run → "needs setup" warning
        _needs_batch_setup = False
        _batch_pending_first_run = False
        _last_batch = ""
        _batch_installed = False
        if _is_admin:
            try:
                _settings = lj(STK, "{}")
                _last_batch = _settings.get("lastBatchRun", "")
                # Check if MorningBatch references us
                try:
                    _mb_txt = model.PythonContent("MorningBatch") or ""
                    if _BATCH_MARKER_START in _mb_txt:
                        _batch_installed = True
                    else:
                        _sn_chk = get_script_name()
                        if ('CallScript("' + _sn_chk + '")' in _mb_txt) or ("CallScript('" + _sn_chk + "')" in _mb_txt):
                            _batch_installed = True
                except:
                    pass
                if _last_batch:
                    try:
                        _lb_dt = datetime.datetime.strptime(_last_batch, "%Y-%m-%d %H:%M:%S")
                        _hours = (datetime.datetime.now() - _lb_dt).total_seconds() / 3600.0
                        if _hours > 36:
                            _needs_batch_setup = True
                    except:
                        _needs_batch_setup = True
                else:
                    # Never run
                    if _batch_installed:
                        _batch_pending_first_run = True  # info only
                    else:
                        _needs_batch_setup = True  # warning
            except:
                pass

        # --- HTML --------------------------------------------------------
        html = '''
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/ag-grid-community@32.0.2/styles/ag-grid.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/ag-grid-community@32.0.2/styles/ag-theme-alpine.css" />
    <style>
    *{box-sizing:border-box}
    .oc{max-width:1100px;margin:0 auto;font-family:Segoe UI,-apple-system,sans-serif}
    .oc-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-wrap:wrap;gap:10px}
    .oc-head h2{margin:0;font-size:22px;color:#333}
    .oc-sub{font-size:12px;color:#888;margin-top:2px}
    .bc{display:flex;align-items:center;gap:6px;margin-bottom:16px;font-size:13px;color:#666}
    .bc a{color:#0078d4;text-decoration:none;cursor:pointer;font-weight:600}
    .bc a:hover{text-decoration:underline}
    .bc .sep{color:#ccc}
    .grd{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:14px;margin-bottom:20px}
    .gc{border:1px solid #e0e0e0;border-radius:12px;background:#fff;overflow:hidden;cursor:pointer;transition:box-shadow .2s,transform .1s}
    .gc:hover{box-shadow:0 4px 16px rgba(0,0,0,.1);transform:translateY(-2px)}
    .gc-top{padding:16px 16px 12px;display:flex;gap:12px;align-items:flex-start}
    .gc-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0}
    .gc-info{flex:1;min-width:0}
    .gc-nm{font-size:15px;font-weight:700;color:#333;margin-bottom:2px}
    .gc-desc{font-size:11px;color:#888;line-height:1.3}
    .gc-bar{height:4px;background:#e8e8e8;margin:0 16px}
    .gc-fill{height:4px;border-radius:2px;transition:width .3s}
    .gc-bot{padding:10px 16px;display:flex;justify-content:space-between;align-items:center;font-size:11px;color:#888}
    .gc-stat{display:flex;gap:12px}
    .gc-s{display:flex;align-items:center;gap:3px}
    .gc-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
    .gc-badge{display:inline-block;padding:1px 6px;border-radius:8px;font-size:10px;font-weight:600;background:#e8f0fd;color:#0078d4;margin-left:6px}
    .gc-acts{display:flex;gap:4px}
    .gc-acts button{padding:3px 8px;border:none;border-radius:4px;font-size:11px;cursor:pointer;background:#f0f0f0;color:#666}
    .gc-acts button:hover{background:#e0e0e0}
    .gc-acts .del{color:#d13438}
    .sum{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:18px}
    .sm{padding:10px 18px;border-radius:10px;text-align:center;min-width:80px}
    .sm .n{font-size:22px;font-weight:700}
    .sm .l{font-size:10px;color:#666;margin-top:2px}
    .freq-tabs{display:flex;border-bottom:2px solid #e0e0e0;margin-bottom:14px;flex-wrap:wrap}
    .ft{padding:9px 18px;cursor:pointer;font-size:13px;font-weight:600;color:#888;border-bottom:2px solid transparent;margin-bottom:-2px}
    .ft.a{color:#0078d4;border-bottom-color:#0078d4}
    .cat-hdr{font-size:13px;font-weight:700;color:#444;margin:14px 0 6px;border-bottom:1px solid #eee;padding-bottom:4px;display:flex;align-items:center;gap:6px}
    .cat-cnt{font-size:11px;font-weight:400;color:#aaa}
    .chk{border:1px solid #e8e8e8;border-radius:8px;margin-bottom:6px;background:#fff;overflow:hidden}
    .rw{display:flex;align-items:center;padding:10px 12px;gap:10px;cursor:pointer}
    .rw:hover{background:#f8f9fa}
    .ic{width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0}
    .ic.g{background:#e8fde8;color:#107c10}
    .ic.y{background:#fff4ce;color:#7a6400}
    .ic.r{background:#fde8e8;color:#d13438}
    .ic.x{background:#f0f0f0;color:#888}
    .inf{flex:1;min-width:0}
    .nm{font-size:13px;font-weight:600;color:#333}
    .mt{font-size:11px;color:#999;margin-top:1px}
    .cnt{font-size:17px;font-weight:700;min-width:40px;text-align:center}
    .cnt.z{color:#107c10}.cnt.w{color:#7a6400}.cnt.c{color:#d13438}
    .acts{display:flex;gap:5px;flex-shrink:0}
    .b{padding:5px 12px;border:none;border-radius:5px;font-size:11px;font-weight:600;cursor:pointer}
    .bp{background:#0078d4;color:#fff}
    .bs{background:#107c10;color:#fff}
    .bo{background:#fff;color:#0078d4;border:1px solid #0078d4}
    .bd{background:#fde8e8;color:#d13438}
    .bt{background:#7c3aed;color:#fff}
    .bg{display:inline-block;padding:1px 6px;border-radius:8px;font-size:10px;font-weight:600;margin-left:4px}
    .bf{background:#e8f0fd;color:#0078d4}
    .ba{background:#e8fde8;color:#107c10}
    .bb{background:#fff4ce;color:#7a6400}
    .bc2{background:#f3e8fd;color:#7c3aed}
    .det{display:none;padding:0 12px 12px;border-top:1px solid #f0f0f0}
    .det.op{display:block}
    .ins{font-size:12px;color:#666;padding:8px;background:#f8f9fa;border-radius:6px;margin-bottom:8px}
    .hist-hdr{font-size:11px;font-weight:700;color:#666;margin:12px 0 5px;padding-top:8px;border-top:1px dashed #eee;display:flex;justify-content:space-between;align-items:center}
    .hist-cnt{font-weight:400;color:#aaa;font-size:10px}
    .hist-row{padding:5px 9px;background:#f8f9fa;border-left:2px solid #e0e0e0;border-radius:0 4px 4px 0;margin-bottom:3px;font-size:11px}
    .hist-by{font-weight:600;color:#333}
    .hist-at{color:#888;margin-left:6px}
    .hist-notes{color:#555;margin-top:2px;font-style:italic}
    .stp{display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f8f8f8}
    .stp:last-child{border-bottom:none}
    .stp input[type=checkbox]{margin:0;flex-shrink:0}
    .sl{font-size:12px;color:#333;flex:0 0 auto;min-width:160px;max-width:45%;line-height:1.3}
    .snp{flex:1;min-width:0;padding:4px 8px;border:1px solid #ddd;border-radius:4px;font-size:11px;margin:0;background:#fff}
    .snp:focus{border-color:#0078d4;outline:none}
    .sn-meta{font-size:10px;color:#aaa;flex-shrink:0;font-style:italic;white-space:nowrap}
    @media (max-width:520px){.sl{min-width:120px;max-width:55%}.sn-meta{display:none}}
    .prg{height:4px;background:#e8e8e8;border-radius:2px;margin:8px 0 4px}
    .prb{height:4px;background:#107c10;border-radius:2px;transition:width .3s}
    .le{padding:7px 10px;border-left:3px solid #107c10;background:#f8f9fa;margin-bottom:5px;border-radius:0 5px 5px 0;font-size:12px}
    .md{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);z-index:9999;justify-content:center;align-items:center}
    .md.on{display:flex}
    .mc{background:#fff;border-radius:12px;padding:22px;max-width:500px;width:90%;max-height:80vh;overflow-y:auto}
    .mc.wide{max-width:700px}
    .fg{margin-bottom:10px}
    .fg label{display:block;font-size:12px;font-weight:600;color:#555;margin-bottom:3px}
    .fg input,.fg select,.fg textarea{width:100%;padding:7px 9px;border:1px solid #ccc;border-radius:5px;font-size:12px}
    .fg textarea{min-height:70px;resize:vertical}
    .ic-pick{display:flex;flex-wrap:wrap;gap:6px;margin-top:4px}
    .ic-opt{width:36px;height:36px;border:2px solid #e0e0e0;border-radius:8px;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:18px}
    .ic-opt.sel{border-color:#0078d4;background:#e8f0fd}
    .clr-pick{display:flex;gap:6px;margin-top:4px}
    .clr-opt{width:28px;height:28px;border-radius:50%;cursor:pointer;border:2px solid transparent}
    .clr-opt.sel{border-color:#333;box-shadow:0 0 0 2px #fff,0 0 0 4px #333}
    /* Catalog */
    .ctlg-cat{margin-bottom:14px}
    .ctlg-cat h4{font-size:13px;font-weight:700;color:#555;margin:0 0 6px;padding-bottom:4px;border-bottom:1px solid #eee}
    .ctlg-item{display:flex;align-items:flex-start;gap:8px;padding:6px 8px;border-radius:6px;cursor:pointer}
    .ctlg-item:hover{background:#f5f5f5}
    .ctlg-item input[type=checkbox]{margin-top:3px;flex-shrink:0}
    .ctlg-nm{font-size:12px;font-weight:600;color:#333}
    .ctlg-desc{font-size:11px;color:#888;margin-top:1px}
    .ctlg-badges{margin-top:2px}
    .ctlg-filter{display:flex;gap:6px;margin-bottom:12px;flex-wrap:wrap}
    .ctlg-fb{padding:4px 10px;border-radius:12px;font-size:11px;font-weight:600;cursor:pointer;border:1px solid #ddd;background:#fff;color:#666}
    .ctlg-fb.active{background:#0078d4;color:#fff;border-color:#0078d4}
    /* Help button + modal */
    .help-btn{display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border:1px solid #0078d4;background:#fff;color:#0078d4;border-radius:50%;font-size:11px;font-weight:700;cursor:pointer;padding:0;margin-left:6px;line-height:1;flex-shrink:0;vertical-align:middle}
    .help-btn:hover{background:#0078d4;color:#fff}
    .help-btn-lg{width:auto;height:auto;border-radius:6px;padding:5px 12px;font-size:11px;gap:4px}
    .help-body{font-size:13px;line-height:1.55;color:#333}
    .help-body h4{font-size:14px;color:#0078d4;margin:14px 0 6px;border-bottom:1px solid #eee;padding-bottom:3px}
    .help-body h4:first-child{margin-top:0}
    .help-body p{margin:6px 0}
    .help-body ul{margin:4px 0 8px 0;padding-left:22px}
    .help-body li{margin:3px 0}
    .help-body code{background:#f0f0f0;padding:1px 5px;border-radius:3px;font-size:11px;font-family:Consolas,Monaco,monospace}
    .help-body .help-tip{background:#f0f7ff;border-left:3px solid #0078d4;padding:8px 10px;margin:8px 0;border-radius:0 4px 4px 0;font-size:12px}
    .help-body .help-warn{background:#fff4ce;border-left:3px solid #f5c842;padding:8px 10px;margin:8px 0;border-radius:0 4px 4px 0;font-size:12px;color:#7a6400}
    .help-toc{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px}
    .help-toc-item{display:block;padding:10px 12px;background:#f8f9fa;border:1px solid #e0e0e0;border-radius:6px;text-decoration:none;color:#333;cursor:pointer}
    .help-toc-item:hover{background:#f0f7ff;border-color:#0078d4}
    .help-toc-nm{font-weight:600;font-size:13px;color:#0078d4}
    .help-toc-desc{font-size:11px;color:#666;margin-top:2px}
    /* Person picker */
    .pp-wrap{position:relative;border:1px solid #ccc;border-radius:5px;padding:3px 6px;min-height:30px;display:flex;flex-wrap:wrap;gap:4px;align-items:center;background:#fff}
    .pp-wrap.focus{border-color:#0078d4}
    .pp-chip{display:inline-flex;align-items:center;gap:5px;background:#e8f0fd;color:#0078d4;padding:2px 8px;border-radius:12px;font-size:11px;line-height:1.4}
    .pp-chip-em{color:#888;font-size:10px;margin-left:3px}
    .pp-chip-x{cursor:pointer;color:#888;font-weight:700;font-size:13px;line-height:1;padding:0 2px}
    .pp-chip-x:hover{color:#d13438}
    .pp-input{border:none;outline:none;flex:1;min-width:120px;font-size:12px;padding:3px;background:transparent}
    .pp-dropdown{position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid #ccc;border-top:none;border-radius:0 0 5px 5px;max-height:240px;overflow-y:auto;z-index:9999;display:none;box-shadow:0 4px 12px rgba(0,0,0,.08)}
    .pp-dropdown.on{display:block}
    .pp-row{padding:6px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0}
    .pp-row:last-child{border-bottom:none}
    .pp-row:hover,.pp-row.active{background:#f0f7ff}
    .pp-row-nm{font-weight:600;color:#333;font-size:12px}
    .pp-row-em{font-size:11px;color:#666;margin-top:1px}
    .pp-row-noem{font-size:11px;color:#d13438;margin-top:1px;font-style:italic}
    .pp-empty{padding:12px;text-align:center;color:#888;font-size:11px}
    /* Toggle switch */
    .tgl{position:relative;display:inline-block;width:36px;height:20px;flex-shrink:0}
    .tgl input{opacity:0;width:0;height:0}
    .tgl .slider{position:absolute;cursor:pointer;top:0;left:0;right:0;bottom:0;background:#ccc;border-radius:20px;transition:.2s}
    .tgl .slider:before{position:absolute;content:"";height:14px;width:14px;left:3px;bottom:3px;background:#fff;border-radius:50%;transition:.2s}
    .tgl input:checked+.slider{background:#107c10}
    .tgl input:checked+.slider:before{transform:translateX(16px)}
    </style>

    <div class="oc">
    <div id="v_groups"></div>
    <div id="v_detail" style="display:none"></div>
    </div>

    <!-- Group Modal -->
    <div id="gm" class="md" onclick="if(event.target===this)this.classList.remove('on')">
    <div class="mc">
        <h3 id="gm_title" style="margin-bottom:14px">Add Group</h3>
        <input type="hidden" id="gm_id">
        <div class="fg"><label>Name</label><input id="gm_name"></div>
        <div class="fg"><label>Description</label><textarea id="gm_desc" style="min-height:50px"></textarea></div>
        <div class="fg"><label>Owner</label><input id="gm_owner" placeholder="Optional"></div>
        <div class="fg"><label>Icon</label>
            <div class="ic-pick" id="gm_icons"></div>
        </div>
        <div class="fg"><label>Color</label>
            <div class="clr-pick" id="gm_colors"></div>
        </div>
        <div style="display:flex;gap:6px;justify-content:flex-end;margin-top:14px">
            <button class="b bo" onclick="document.getElementById('gm').classList.remove('on')">Cancel</button>
            <button class="b bp" onclick="saveGroup()">Save</button>
        </div>
    </div></div>

    <!-- Check Modal -->
    <div id="ckm" class="md" onclick="if(event.target===this)this.classList.remove('on')">
    <div class="mc">
        <h3 style="margin-bottom:14px"><span id="ckm_title">Add Check</span> <button class="help-btn" onclick="showHelp('check_types')" title="Check types & how to fill this in">?</button></h3>
        <input type="hidden" id="ckm_id">
        <input type="hidden" id="ckm_gid">
        <div class="fg"><label>Name</label><input id="ckm_name"></div>
        <div class="fg"><label>Category</label><select id="ckm_cat"><option>People</option><option>Involvements</option><option>Email</option><option>Finance</option><option>Facilities</option><option>General</option><option>Archive</option></select></div>
        <div class="fg"><label>Frequency</label><select id="ckm_freq" onchange="onFreqChange()"><option value="daily">Daily</option><option value="weekly">Weekly</option><option value="monthly" selected>Monthly</option><option value="annual">Annual</option></select></div>
        <div id="ckm_due_wrap" style="display:none">
            <div id="ckm_due_daily" style="display:none" class="fg">
                <label>Active Days</label>
                <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center">
                    <label style="display:flex;align-items:center;gap:3px;font-weight:400;font-size:12px"><input type="checkbox" id="ckm_dow_0" style="width:auto"> Mon</label>
                    <label style="display:flex;align-items:center;gap:3px;font-weight:400;font-size:12px"><input type="checkbox" id="ckm_dow_1" style="width:auto"> Tue</label>
                    <label style="display:flex;align-items:center;gap:3px;font-weight:400;font-size:12px"><input type="checkbox" id="ckm_dow_2" style="width:auto"> Wed</label>
                    <label style="display:flex;align-items:center;gap:3px;font-weight:400;font-size:12px"><input type="checkbox" id="ckm_dow_3" style="width:auto"> Thu</label>
                    <label style="display:flex;align-items:center;gap:3px;font-weight:400;font-size:12px"><input type="checkbox" id="ckm_dow_4" style="width:auto"> Fri</label>
                    <label style="display:flex;align-items:center;gap:3px;font-weight:400;font-size:12px"><input type="checkbox" id="ckm_dow_5" style="width:auto"> Sat</label>
                    <label style="display:flex;align-items:center;gap:3px;font-weight:400;font-size:12px"><input type="checkbox" id="ckm_dow_6" style="width:auto"> Sun</label>
                    <span style="margin-left:6px;display:flex;gap:4px">
                        <button type="button" class="b bo" style="font-size:10px;padding:2px 6px" onclick="setDows('weekdays')">Weekdays</button>
                        <button type="button" class="b bo" style="font-size:10px;padding:2px 6px" onclick="setDows('all')">All</button>
                        <button type="button" class="b bo" style="font-size:10px;padding:2px 6px" onclick="setDows('weekends')">Weekends</button>
                    </span>
                </div>
                <p style="font-size:11px;color:#888;margin:3px 0 0">Reminders only fire on the days you check. Default for new daily checks is weekdays.</p>
            </div>
            <div id="ckm_due_weekly" style="display:none" class="fg"><label>Due Day</label><select id="ckm_due_day_w"><option value="0">Monday</option><option value="1">Tuesday</option><option value="2">Wednesday</option><option value="3">Thursday</option><option value="4">Friday</option><option value="5">Saturday</option><option value="6">Sunday</option></select></div>
            <div id="ckm_due_monthly" style="display:none" class="fg"><label>Due Day of Month</label><select id="ckm_due_day_m"><option value="1">1st</option><option value="5">5th</option><option value="10">10th</option><option value="15">15th</option><option value="20">20th</option><option value="25">25th</option><option value="28">28th</option></select></div>
            <div id="ckm_due_annual" style="display:none;gap:6px">
                <div class="fg" style="flex:1"><label>Due Month</label><select id="ckm_due_month_a"><option value="1">January</option><option value="2">February</option><option value="3">March</option><option value="4">April</option><option value="5">May</option><option value="6">June</option><option value="7">July</option><option value="8">August</option><option value="9">September</option><option value="10">October</option><option value="11">November</option><option value="12">December</option></select></div>
                <div class="fg" style="flex:1"><label>Due Day</label><select id="ckm_due_day_a"><option value="1">1st</option><option value="5">5th</option><option value="10">10th</option><option value="15">15th</option><option value="20">20th</option><option value="25">25th</option><option value="28">28th</option></select></div>
            </div>
        </div>
        <div class="fg"><label>Type</label><select id="ckm_type" onchange="onTypeChange()"><option value="manual">Manual</option><option value="auto">Automated (SQL)</option><option value="search">Saved Search</option></select></div>
        <div id="ckm_sql_wrap" style="display:none">
            <div class="fg"><label>SQL (return column named cnt)</label><textarea id="ckm_sql"></textarea></div>
            <div style="display:flex;gap:6px">
                <div class="fg" style="flex:1"><label>Warn threshold</label><input type="number" id="ckm_th" value="0"></div>
                <div class="fg" style="flex:1"><label>Critical threshold</label><input type="number" id="ckm_cr" value="5"></div>
            </div>
        </div>
        <div id="ckm_search_wrap" style="display:none">
            <div class="fg"><label>Search Code or Saved Search Name</label>
                <div style="display:flex;gap:6px;position:relative">
                    <div style="flex:1;position:relative">
                        <input id="ckm_search" placeholder="Type to search saved searches, or paste a condition code" style="width:100%" autocomplete="off">
                        <div id="ckm_search_dd" class="pp-dropdown" style="display:none"></div>
                    </div>
                    <button type="button" class="b bp" onclick="lookupSearch()" style="white-space:nowrap">Lookup</button>
                </div>
            </div>
            <div id="ckm_search_result" style="display:none;margin:6px 0 8px;padding:10px;background:#f8f9fa;border-radius:6px;border:1px solid #e0e0e0;font-size:12px"></div>
            <p style="font-size:11px;color:#888;margin:0 0 6px">Enter a condition code or saved search name. Click Lookup to view a saved search's definition.</p>
            <div style="display:flex;gap:6px">
                <div class="fg" style="flex:1"><label>Warn threshold</label><input type="number" id="ckm_th2" value="0"></div>
                <div class="fg" style="flex:1"><label>Critical threshold</label><input type="number" id="ckm_cr2" value="5"></div>
            </div>
        </div>
        <div class="fg"><label>Instructions</label><textarea id="ckm_inst"></textarea></div>
        <div class="fg"><label>Steps (one per line)</label><textarea id="ckm_steps" placeholder="Step 1\nStep 2"></textarea></div>
        <div class="fg"><label>Email Reminders <span style="font-weight:400;color:#888;font-size:11px">(when due)</span></label>
            <div id="ckm_notify_picker"></div>
            <input type="hidden" id="ckm_notify_pids">
            <p style="font-size:11px;color:#888;margin:2px 0 0">Picked people get the morning-batch reminder when this check is due. Leave empty to use the Default Recipient from Settings.</p>
        </div>
        <div class="fg"><label>Completion Notifications <span style="font-weight:400;color:#888;font-size:11px">(when completed)</span></label>
            <div id="ckm_complete_notify_picker"></div>
            <input type="hidden" id="ckm_complete_notify_pids">
            <p style="font-size:11px;color:#888;margin:2px 0 0">Picked people get an email when someone marks this check complete &mdash; useful for ministry leaders or supervisors who need to be notified.</p>
        </div>
        <div style="display:flex;gap:6px;justify-content:flex-end;margin-top:14px">
            <button class="b bo" onclick="document.getElementById('ckm').classList.remove('on')">Cancel</button>
            <button class="b bp" onclick="saveCheck()">Save</button>
        </div>
    </div></div>

    <!-- Catalog Modal -->
    <div id="catm" class="md" onclick="if(event.target===this)this.classList.remove('on')">
    <div class="mc wide">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:4px">
            <h3 style="margin:0">Check Library <button class="help-btn" onclick="showHelp('manage_library')" title="About the marketplace">?</button></h3>
            <button class="b bt" onclick="document.getElementById('catm').classList.remove('on');document.getElementById('subm').classList.add('on')" style="font-size:11px">Share a Check</button>
        </div>
        <p style="font-size:12px;color:#888;margin:0 0 12px">Toggle checks on/off for this group. Items marked <span style="background:#f3e8fd;color:#7c3aed;padding:1px 5px;border-radius:6px;font-size:10px;font-weight:600">community</span> were shared by other churches.</p>
        <div id="ctlg_update_all_bar" style="display:none;background:#fff4ce;border:1px solid #f5c842;border-radius:6px;padding:8px 12px;margin-bottom:10px;display:flex;justify-content:space-between;align-items:center;gap:10px">
            <div style="font-size:12px;color:#7a6400"><strong id="ctlg_update_all_count">0</strong> installed check(s) in this group have updates available from the marketplace.</div>
            <button class="b bp" onclick="updateAllMarketplace(curGroup)" style="white-space:nowrap;font-size:11px">Update All</button>
        </div>
        <div class="ctlg-filter" id="ctlg_filter"></div>
        <div id="ctlg_list"></div>
        <div style="display:flex;gap:6px;justify-content:space-between;align-items:center;margin-top:14px;border-top:1px solid #eee;padding-top:12px">
            <span id="ctlg_count" style="font-size:12px;color:#888">0 enabled</span>
            <div style="display:flex;gap:6px">
                <button class="b bo" onclick="document.getElementById('catm').classList.remove('on')">Cancel</button>
                <button class="b bt" onclick="importFromCatalog()">Save Changes</button>
            </div>
        </div>
    </div></div>

    <!-- Submit to Marketplace Modal -->
    <div id="subm" class="md" onclick="if(event.target===this)this.classList.remove('on')">
    <div class="mc">
        <h3 style="margin-bottom:4px">Share with the Community</h3>
        <p style="font-size:12px;color:#888;margin:0 0 12px">Submit a check for other churches to use. It will be reviewed before appearing in the marketplace.</p>
        <div class="fg"><label>Name</label><input id="sub_name"></div>
        <div class="fg"><label>Category</label><select id="sub_cat"><option>People</option><option>Involvements</option><option>Email</option><option>Finance</option><option>Facilities</option><option>General</option><option>Archive</option></select></div>
        <div class="fg"><label>Frequency</label><select id="sub_freq"><option value="daily">Daily</option><option value="weekly">Weekly</option><option value="monthly" selected>Monthly</option><option value="annual">Annual</option></select></div>
        <div class="fg"><label>Type</label><select id="sub_type" onchange="document.getElementById('sub_sql_w').style.display=this.value==='auto'?'block':'none';document.getElementById('sub_search_w').style.display=this.value==='search'?'block':'none'"><option value="manual">Manual</option><option value="auto">SQL</option><option value="search">Saved Search</option></select></div>
        <div id="sub_sql_w" style="display:none">
            <div class="fg"><label>SQL (return column named cnt)</label><textarea id="sub_sql"></textarea></div>
        </div>
        <div id="sub_search_w" style="display:none">
            <div class="fg"><label>Search Builder Condition Code</label><input id="sub_search" oninput="checkSearchPortable(this.value)"></div>
            <div id="sub_search_warn" style="display:none;font-size:11px;color:#d13438;margin:4px 0 6px;padding:6px 8px;background:#fde8e8;border-radius:4px">This looks like a saved search name, not a condition code. Saved search names are specific to your church and won't work at other churches. Use condition codes instead.</div>
            <p style="font-size:11px;color:#888;margin:0 0 6px">Use Search Builder condition codes for portability: <code>MemberStatusId=50</code>, <code>IsMemberOnly=1[True] AND HasEmail=0[False]</code>. Saved search names only work at your church.</p>
        </div>
        <div style="display:flex;gap:6px">
            <div class="fg" style="flex:1"><label>Warn threshold</label><input type="number" id="sub_th" value="0"></div>
            <div class="fg" style="flex:1"><label>Critical threshold</label><input type="number" id="sub_cr" value="5"></div>
        </div>
        <div class="fg"><label>Instructions</label><textarea id="sub_inst" style="min-height:50px"></textarea></div>
        <div class="fg"><label>Steps (one per line)</label><textarea id="sub_steps" placeholder="Step 1\nStep 2" style="min-height:50px"></textarea></div>
        <div style="display:flex;gap:6px;justify-content:flex-end;margin-top:14px">
            <button class="b bo" onclick="document.getElementById('subm').classList.remove('on')">Cancel</button>
            <button class="b bt" onclick="submitToMarketplace()">Submit for Review</button>
        </div>
    </div></div>

    <!-- Settings Modal -->
    <div id="setm" class="md" onclick="if(event.target===this)this.classList.remove('on')">
    <div class="mc">
        <h3 style="margin-bottom:14px">Settings <button class="help-btn" onclick="showHelp('toc')" title="All help topics">?</button></h3>

        <h4 style="font-size:13px;color:#444;margin:0 0 10px;border-bottom:1px solid #eee;padding-bottom:4px">Permissions <button class="help-btn" onclick="showHelp('permissions')" title="About permissions">?</button></h4>
        <div class="fg"><label>Edit Roles</label><input id="set_edit_roles" placeholder="Admin">
            <p style="font-size:11px;color:#888;margin:2px 0 0">Comma-separated TouchPoint roles allowed to add, edit, or delete groups and checks. Example: <code>Admin,Staff</code>. Admins always have full access.</p></div>
        <div class="fg"><label>Complete Roles</label><input id="set_complete_roles" placeholder="Admin">
            <p style="font-size:11px;color:#888;margin:2px 0 0">Roles allowed to mark checks complete, toggle steps, and run automated checks. Example: <code>Admin,Staff,Volunteer</code>. Editors can also complete.</p></div>

        <h4 style="font-size:13px;color:#444;margin:18px 0 10px;border-bottom:1px solid #eee;padding-bottom:4px">Reminders</h4>
        <div class="fg"><label>Default Recipient</label>
            <div id="set_default_picker"></div>
            <input type="hidden" id="set_default_recipient_pid">
            <input type="hidden" id="set_email">
            <p style="font-size:11px;color:#888;margin:2px 0 0">Receives reminders for any check that doesn't have specific people assigned. Email is resolved at send time from the person's TouchPoint profile.</p></div>
        <div class="fg" style="display:flex;align-items:center;gap:8px">
            <input type="checkbox" id="set_cc" style="width:auto">
            <label for="set_cc" style="margin:0;cursor:pointer">Also send to default email when specific people are assigned</label>
        </div>

        <h4 style="font-size:13px;color:#444;margin:18px 0 10px;border-bottom:1px solid #eee;padding-bottom:4px">Morning Batch <button class="help-btn" onclick="showHelp('morning_batch')" title="About morning batch">?</button> <span id="set_batch_status" style="font-size:11px;font-weight:400"></span></h4>
        <div style="display:flex;align-items:center;gap:10px;padding:8px 10px;margin-bottom:10px;background:#f0f7ff;border:1px solid #cfe3ff;border-radius:6px">
            <label class="tgl"><input type="checkbox" id="set_batch_enabled" checked><span class="slider"></span></label>
            <div style="flex:1">
                <label for="set_batch_enabled" style="margin:0;cursor:pointer;font-size:13px;font-weight:600;color:#333">Send reminder emails when morning batch fires</label>
                <p style="margin:2px 0 0;font-size:11px;color:#666">Toggle off to pause automatic sends without removing the wrapper from your morning batch &mdash; safe when other scripts share the batch.</p>
            </div>
        </div>
        <div id="set_install_panel" style="margin:0 0 10px;padding:12px;background:#f8f9fa;border:1px solid #e0e0e0;border-radius:6px">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap">
                <div style="font-size:13px;color:#333"><strong>One-click install</strong> &mdash; let this script add itself to your TouchPoint <code>MorningBatch</code> content. Existing scripts in MorningBatch are left untouched.</div>
                <div style="display:flex;gap:6px">
                    <button id="install_batch_btn" type="button" class="b bp" onclick="installBatch()" style="white-space:nowrap">&#10004; Install</button>
                    <button id="uninstall_batch_btn" type="button" class="b bd" onclick="uninstallBatch()" style="white-space:nowrap;display:none">&#10005; Remove</button>
                </div>
            </div>
            <div id="install_batch_status" style="margin-top:8px;font-size:12px;color:#888"></div>
            <div id="install_batch_result" style="display:none;margin-top:8px;padding:8px 10px;border-radius:4px;font-size:12px"></div>
        </div>
        <details style="margin:0 0 10px">
            <summary style="cursor:pointer;font-size:12px;color:#666;padding:4px 0">Manual setup (if you prefer to edit MorningBatch yourself)</summary>
            <div style="margin:8px 0 0;padding:10px;background:#f8f9fa;border-radius:6px;font-size:12px;color:#666">
                <p style="margin:0 0 8px">Add this block to your existing <code>MorningBatch</code> Python special content:</p>
                <pre id="set_wrapper_code" style="margin:0;padding:8px;background:#fff;border:1px solid #e0e0e0;border-radius:4px;font-size:11px;overflow-x:auto;white-space:pre">Data.run_batch = "true"
model.CallScript("TPxi_OpsChecklists")</pre>
                <p style="margin:8px 0 0">Calls this script (currently <code id="set_script_name">TPxi_OpsChecklists</code>) which sends one consolidated reminder email per person for items due that day.</p>
                <p style="margin:8px 0 0;display:none" id="set_wrapper_legacy">Or as a separate wrapper named <code id="set_wrapper_name">TPxi_OpsChecklistsBatch</code> if you want to call it from another scheduler.</p>
            </div>
        </details>
        <div style="margin:10px 0;padding:10px;background:#f0f7ff;border:1px solid #cfe3ff;border-radius:6px">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap">
                <div style="font-size:12px;color:#333"><strong>Send reminders now</strong> &mdash; trigger today's reminder emails immediately without waiting for morning batch.</div>
                <button id="send_now_btn" type="button" class="b bp" onclick="sendRemindersNow()" style="white-space:nowrap">&#9993; Send Now</button>
            </div>
            <div id="send_now_result" style="display:none;margin-top:8px;padding:8px 10px;border-radius:4px;font-size:12px"></div>
            <div id="send_now_last" style="font-size:11px;color:#888;margin-top:6px"></div>
        </div>
        <div id="set_admin_warn" style="display:none;margin:8px 0;padding:8px 10px;background:#fff4ce;border-radius:6px;font-size:12px;color:#7a6400">Only TouchPoint Admins can save settings.</div>
        <div style="display:flex;gap:6px;justify-content:flex-end;margin-top:14px">
            <button class="b bo" onclick="document.getElementById('setm').classList.remove('on')">Cancel</button>
            <button id="set_save_btn" class="b bp" onclick="saveSettings()">Save</button>
        </div>
    </div></div>

    <!-- Help Modal -->
    <div id="hm" class="md" onclick="if(event.target===this)this.classList.remove('on')">
    <div class="mc" style="max-width:720px">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px;margin-bottom:12px;border-bottom:1px solid #eee;padding-bottom:10px">
            <div style="flex:1">
                <h3 id="hm_title" style="margin:0;font-size:17px;color:#0078d4">Help</h3>
                <div id="hm_breadcrumb" style="font-size:11px;color:#888;margin-top:2px"></div>
            </div>
            <button class="b bo" onclick="document.getElementById('hm').classList.remove('on')" style="font-size:11px;padding:3px 10px">Close</button>
        </div>
        <div id="hm_body" class="help-body"></div>
        <div id="hm_footer" style="margin-top:14px;padding-top:10px;border-top:1px solid #eee;display:flex;justify-content:space-between;align-items:center;font-size:11px;color:#888">
            <a onclick="showHelp('toc')" style="color:#0078d4;cursor:pointer">&#8617; Back to all topics</a>
        </div>
    </div></div>

    <!-- SQL Detail Viewer Modal (ag-grid) -->
    <div id="dvm" class="md" onclick="if(event.target===this)this.classList.remove('on')">
    <div class="mc" style="max-width:1100px;width:95%;max-height:none;overflow:visible">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px;margin-bottom:10px">
            <div style="flex:1">
                <h3 style="margin:0;font-size:16px"><span id="dvm_title">Check Details</span> <button class="help-btn" onclick="showHelp('view_drilldown')" title="Drilldown features">?</button></h3>
                <div id="dvm_meta" style="font-size:11px;color:#888;margin-top:2px"></div>
            </div>
            <div style="display:flex;gap:6px">
                <button class="b bo" onclick="exportGridCsv()" style="font-size:11px">&#8595; CSV</button>
                <button class="b bo" onclick="document.getElementById('dvm').classList.remove('on')" style="font-size:11px">Close</button>
            </div>
        </div>
        <div id="dvm_action_bar" style="display:none;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px;padding:8px 12px;margin-bottom:8px;align-items:center;gap:10px;font-size:12px">
            <span id="dvm_sel_count" style="font-weight:600;color:#166534"></span>
            <button class="b bp" onclick="openTagModal()" style="font-size:11px">&#127991;&#65039; Tag Selected</button>
            <button class="b bo" onclick="clearGridSelection()" style="font-size:11px">Clear</button>
        </div>
        <div id="dvm_grid" class="ag-theme-alpine" style="width:100%;height:500px"></div>
        <div id="dvm_status" style="font-size:11px;color:#888;margin-top:6px"></div>
    </div></div>

    <!-- Bulk Tag Modal -->
    <div id="tagm" class="md" onclick="if(event.target===this)this.classList.remove('on')">
    <div class="mc">
        <h3 style="margin-bottom:6px">Tag Selected People</h3>
        <p id="tagm_count" style="font-size:12px;color:#888;margin:0 0 10px"></p>
        <div class="fg">
            <label>Tag Name</label>
            <input id="tagm_name" placeholder="e.g. FollowUpNeeded">
            <p style="font-size:11px;color:#888;margin:2px 0 0">A new tag is created (or existing tag is added to). Owned by you.</p>
        </div>
        <div style="display:flex;gap:6px;justify-content:flex-end;margin-top:14px">
            <button class="b bo" onclick="document.getElementById('tagm').classList.remove('on')">Cancel</button>
            <button id="tagm_submit_btn" class="b bp" onclick="submitBulkTag()">Tag</button>
        </div>
    </div></div>

    <!-- Complete Check Modal -->
    <div id="cm" class="md" onclick="if(event.target===this)this.classList.remove('on')">
    <div class="mc">
        <h3 style="margin-bottom:14px">Mark Complete</h3>
        <p id="cm_name" style="font-size:13px;color:#555;margin-bottom:10px"></p>
        <div class="fg"><label>Notes</label><textarea id="cm_notes" placeholder="Summary of what was done..."></textarea></div>
        <div style="display:flex;gap:6px;justify-content:flex-end">
            <button class="b bo" onclick="document.getElementById('cm').classList.remove('on')">Cancel</button>
            <button class="b bs" onclick="completeCheck()">Complete</button>
        </div>
    </div></div>
    '''

        model.Form = html

        # --- JavaScript --------------------------------------------------
        js = """
    var G = JSON.parse('""" + gj + """');
    var AC = JSON.parse('""" + acj + """');
    var R = JSON.parse('""" + rj + """');
    var LD = JSON.parse('""" + ldj + """');
    var LC = JSON.parse('""" + lcj + """');
    var SD = JSON.parse('""" + sdj + """');
    var GS = JSON.parse('""" + gsj + """');
    var LOG = JSON.parse('""" + logj + """');
    var CAN_EDIT = """ + ("true" if _can_edit else "false") + """;
    var CAN_COMPLETE = """ + ("true" if _can_complete else "false") + """;
    var IS_ADMIN = """ + ("true" if _is_admin else "false") + """;
    var NEEDS_BATCH_SETUP = """ + ("true" if _needs_batch_setup else "false") + """;
    var BATCH_INSTALLED = """ + ("true" if _batch_installed else "false") + """;
    var BATCH_PENDING_FIRST_RUN = """ + ("true" if _batch_pending_first_run else "false") + """;
    var LAST_BATCH_RUN = '""" + str(_last_batch).replace("'", "\\'") + """';
    var SCRIPT_NAME = '""" + str(get_script_name()).replace("'", "\\'") + """';
    var APP_VERSION = '""" + APP_VERSION + """';
    var DC_API_BASE = '""" + DC_API_BASE + """';
    var DC_SCRIPT_ID = '""" + DC_SCRIPT_ID + """';
    // Override SCRIPT_NAME from the live URL — most reliable source.
    (function() {
        try {
            var p = window.location.pathname.split('/');
            for (var i = 0; i < p.length; i++) {
                if ((p[i] === 'PyScript' || p[i] === 'PyScriptForm') && i + 1 < p.length) {
                    var nm = decodeURIComponent(p[i+1]).split('?')[0];
                    if (nm) SCRIPT_NAME = nm;
                    break;
                }
            }
        } catch(e) {}
    })();
    // Inject SCRIPT_NAME into every $.post so the server can persist it for
    // morning-batch use (which has no URL of its own).
    (function() {
        if (window.$ && $.ajaxPrefilter) {
            $.ajaxPrefilter(function(opts) {
                if (opts.type && opts.type.toUpperCase() === 'POST' && opts.url === window.location.pathname) {
                    if (typeof opts.data === 'string') {
                        if (opts.data.indexOf('script_name=') === -1) {
                            opts.data += '&script_name=' + encodeURIComponent(SCRIPT_NAME || '');
                        }
                    } else if (opts.data && typeof opts.data === 'object' && !opts.data.script_name) {
                        opts.data.script_name = SCRIPT_NAME || '';
                    }
                }
            });
        }
    })();
    function esc(s) {
        if (s === undefined || s === null) return '';
        return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
    }
    function attr(s) { return esc(s); }

    // ============================================================
    //  HELP CONTENT — keyed by topic id; rendered into Help modal
    // ============================================================
    var HELP_TOPICS = {
        toc: {
            title: 'Help & Tutorials',
            crumb: '',
            body: '<p>Click any topic below for guidance. You can also click the <span class="help-btn" style="cursor:default">?</span> icon next to any feature for context-specific help.</p>' +
                '<div class="help-toc">' +
                '<a class="help-toc-item" onclick="showHelp(\\'getting_started\\')"><div class="help-toc-nm">&#9889; Getting Started</div><div class="help-toc-desc">First time? Start here.</div></a>' +
                '<a class="help-toc-item" onclick="showHelp(\\'dashboard\\')"><div class="help-toc-nm">&#127968; Groups Dashboard</div><div class="help-toc-desc">The main view: groups, status counts, recent activity.</div></a>' +
                '<a class="help-toc-item" onclick="showHelp(\\'group_detail\\')"><div class="help-toc-nm">&#128203; Group Detail</div><div class="help-toc-desc">Tabs, categories, running checks.</div></a>' +
                '<a class="help-toc-item" onclick="showHelp(\\'check_types\\')"><div class="help-toc-nm">&#9881; Check Types</div><div class="help-toc-desc">Manual / SQL / Saved Search explained.</div></a>' +
                '<a class="help-toc-item" onclick="showHelp(\\'frequencies\\')"><div class="help-toc-nm">&#128197; Frequencies</div><div class="help-toc-desc">Daily, weekly, monthly, annual scheduling.</div></a>' +
                '<a class="help-toc-item" onclick="showHelp(\\'manage_library\\')"><div class="help-toc-nm">&#128218; Manage Library</div><div class="help-toc-desc">Browse marketplace, install, update community checks.</div></a>' +
                '<a class="help-toc-item" onclick="showHelp(\\'view_drilldown\\')"><div class="help-toc-nm">&#128269; View / Drilldown</div><div class="help-toc-desc">ag-grid viewer, links, bulk tagging, CSV export.</div></a>' +
                '<a class="help-toc-item" onclick="showHelp(\\'calendar\\')"><div class="help-toc-nm">&#128197; Calendar</div><div class="help-toc-desc">Per-group month view of due / completed / skipped.</div></a>' +
                '<a class="help-toc-item" onclick="showHelp(\\'morning_batch\\')"><div class="help-toc-nm">&#9203; Morning Batch</div><div class="help-toc-desc">Auto reminders via TouchPoint scheduler.</div></a>' +
                '<a class="help-toc-item" onclick="showHelp(\\'permissions\\')"><div class="help-toc-nm">&#128272; Permissions</div><div class="help-toc-desc">Edit vs Complete vs Admin roles.</div></a>' +
                '</div>'
        },
        getting_started: {
            title: 'Getting Started',
            crumb: 'Help &raquo; Getting Started',
            body: '<h4>What is this?</h4>' +
                '<p>Operations Checklists keeps recurring church operations on track &mdash; from data hygiene checks (members missing email) to weekly Sunday prep (AV tested, doors unlocked) to month-end finance close. Everything lives in <strong>groups</strong> of related <strong>checks</strong>.</p>' +
                '<h4>Three steps to get value in 5 minutes</h4>' +
                '<ol><li><strong>Open Manage Library</strong> in any group &rarr; toggle on a few community checks (e.g. "Members missing email", "Just Addeds review")</li>' +
                '<li><strong>Click "Run All Checks"</strong> &rarr; counts populate the dashboard</li>' +
                '<li><strong>Click View on any check</strong> &rarr; see the actual people who need attention, with links to their profiles</li></ol>' +
                '<h4>Make it automatic</h4>' +
                '<p>Open <strong>Settings</strong> &rarr; <strong>Morning Batch</strong> &rarr; click <strong>Install</strong>. Each morning the checks refresh and a consolidated reminder email goes out to whoever you assigned.</p>' +
                '<div class="help-tip"><strong>Tip:</strong> You can group your checks however makes sense. "Data Health" for SQL audits, "Sunday Operations" for weekly Sunday prep, "Finance Close" for month-end &mdash; or one big "Operations" group if you prefer.</div>'
        },
        dashboard: {
            title: 'Groups Dashboard',
            crumb: 'Help &raquo; Dashboard',
            body: '<h4>What you see</h4>' +
                '<p>The main view shows all your <strong>groups</strong>. Each group card displays:</p>' +
                '<ul><li><strong>Progress bar</strong> &mdash; percent of checks done in current period</li>' +
                '<li><strong>Color dots</strong> &mdash; how many checks are <span style="color:#107c10">passing</span>, <span style="color:#7a6400">warning</span>, or <span style="color:#d13438">critical</span></li>' +
                '<li><strong>X/Y done</strong> &mdash; count completed of total in this group</li></ul>' +
                '<h4>The summary tiles at the top</h4>' +
                '<p>Aggregate counts across all groups: completed, total checks, issues found, group count.</p>' +
                '<h4>Banners</h4>' +
                '<ul><li><strong>Yellow "Morning batch reminders are not active"</strong> &mdash; install via Settings to enable automated daily emails</li>' +
                '<li><strong>Blue "marketplace updates available"</strong> &mdash; new versions of catalog items are out; click Update All to pull them</li></ul>'
        },
        group_detail: {
            title: 'Group Detail',
            crumb: 'Help &raquo; Group Detail',
            body: '<h4>Tabs</h4>' +
                '<ul><li><strong>All / Daily / Weekly / Monthly / Annual</strong> &mdash; filter checks by frequency</li>' +
                '<li><strong>Calendar</strong> &mdash; month-grid view of when each check is due</li>' +
                '<li><strong>History</strong> &mdash; chronological log of completions in this group</li></ul>' +
                '<h4>Action buttons</h4>' +
                '<ul><li><strong>&#9654; Run All Checks</strong> &mdash; execute every SQL/search check in this group; populates counts</li>' +
                '<li><strong>Manage Library</strong> &mdash; browse and install community-shared checks</li>' +
                '<li><strong>+ Add Check</strong> &mdash; create your own (manual, SQL, or saved-search based)</li>' +
                '<li><strong>&#8634; Reset</strong> &mdash; clear current-period progress (history kept)</li></ul>' +
                '<h4>Each check row shows</h4>' +
                '<p>Status icon, name + badges (frequency, type, count), last-run info, and per-row buttons:</p>' +
                '<ul><li><strong>Run</strong> &mdash; refresh just this check</li>' +
                '<li><strong>View (N)</strong> &mdash; opens the drilldown grid (auto/search checks with rows)</li>' +
                '<li><strong>&#9998; Edit</strong>, <strong>&#10003; Complete</strong>, <strong>&#10005; Delete</strong></li></ul>' +
                '<p>Click anywhere on a check row to expand its instructions, steps, per-step notes, and recent completion history.</p>'
        },
        check_types: {
            title: 'Check Types',
            crumb: 'Help &raquo; Check Types',
            body: '<h4>Manual</h4>' +
                '<p>A simple checklist item with steps. No SQL, no automation. Great for "Verify building unlocked", "AV system tested", "Welcome letter sent".</p>' +
                '<h4>Automated (SQL)</h4>' +
                '<p>Runs raw SQL against your TouchPoint database. Two patterns:</p>' +
                '<ul><li><strong>Count-only:</strong> <code>SELECT COUNT(*) as cnt FROM ... WHERE ...</code> &mdash; just shows a number</li>' +
                '<li><strong>Row-returning (recommended):</strong> <code>SELECT TOP 100 PeopleId, Name, ... FROM ... WHERE ...</code> &mdash; count comes from row count, AND the View button opens the ag-grid drilldown showing each person</li></ul>' +
                '<div class="help-tip"><strong>Pro tip:</strong> Always include <code>PeopleId</code> as the first column in row-returning queries. The drilldown viewer makes Name links clickable to <code>/Person2/&lt;id&gt;</code>, and bulk-tagging works.</div>' +
                '<h4>Saved Search</h4>' +
                '<p>References a TouchPoint Search Builder query. Two ways:</p>' +
                '<ul><li><strong>By name:</strong> Type the saved search name (autocomplete will help). View link goes to TouchPoint\\'s full Search Builder UI.</li>' +
                '<li><strong>By condition code:</strong> e.g. <code>MemberStatusId = 50</code>. Works for simple People columns. Multi-condition codes with virtual fields (IsMemberOnly, BirthdayThisWeek) often fail &mdash; save those as named searches instead.</li></ul>' +
                '<div class="help-warn">Search Builder virtual fields (IsMemberOnly, HasEmail, BirthdayThisWeek) and function predicates (RecentVisitCount(...)) don\\'t parse reliably as raw codes. Save them as named searches in TouchPoint Search Builder for reliable counts.</div>'
        },
        frequencies: {
            title: 'Frequencies & Scheduling',
            crumb: 'Help &raquo; Frequencies',
            body: '<h4>Daily</h4>' +
                '<p>Runs every day &mdash; or only on selected days. <strong>Active Days</strong> picker lets you choose: defaults to <strong>Mon-Fri</strong> for new daily checks. Pick "All" for 7-day, "Weekends" for Sat-Sun only, or any custom subset.</p>' +
                '<h4>Weekly</h4>' +
                '<p>Due once a week on the day you pick (Monday default). Period runs from that day to the next.</p>' +
                '<h4>Monthly</h4>' +
                '<p>Due on a specific day of the month (1, 5, 10, 15, 20, 25, or 28). Period spans from that day to the next month\\'s same day.</p>' +
                '<h4>Annual</h4>' +
                '<p>Due once a year on a specific month + day. Useful for "Year-end giving statements" type tasks.</p>' +
                '<h4>How "due" works</h4>' +
                '<p>Each frequency has a <strong>period</strong>. A check is "due" in its current period until completed. Completing it counts for that period; the next period starts fresh. The morning batch only emails reminders for checks that are due AND not yet completed in the current period.</p>'
        },
        manage_library: {
            title: 'Manage Library',
            crumb: 'Help &raquo; Manage Library',
            body: '<h4>What is the Library?</h4>' +
                '<p>A community marketplace of pre-built checks. Browse by category (People / Involvements / Email / Finance / Facilities / General / Archive), toggle on the ones you want, click Save Changes &mdash; they get installed in your group.</p>' +
                '<h4>Badges</h4>' +
                '<ul><li><span class="bg ba">enabled</span> &mdash; already installed in this group</li>' +
                '<li><span class="bg" style="background:#dff6dd;color:#107c10;border:1px solid #92c89e">&#10024; new</span> &mdash; added to the marketplace within the last 14 days and not yet installed</li>' +
                '<li><span class="bg" style="background:#fff4ce;color:#7a6400">modified locally</span> &mdash; you\\'ve edited the SQL/instructions vs. the marketplace original. Updates would overwrite your changes (with warning).</li>' +
                '<li><span class="bg" style="background:#fde8e8;color:#d13438">update available</span> &mdash; marketplace has a newer version. Click the Update button or use Update All.</li>' +
                '<li><span class="bg" style="background:#fef3c7;color:#92400e;border:1px solid #fbbf24">&#11088; default</span> &mdash; curated default item, auto-installed in its target group during initial setup</li>' +
                '<li><span class="bg bc2">community</span> &mdash; submitted by another church (vs. original DisplayCache curation)</li></ul>' +
                '<h4>Update flows</h4>' +
                '<ul><li><strong>Per-item Update</strong> &mdash; each item with a newer rev shows its own Update button</li>' +
                '<li><strong>Update All (in modal)</strong> &mdash; bulk-update everything in this group</li>' +
                '<li><strong>Update All (dashboard banner)</strong> &mdash; org-wide across all your groups</li></ul>' +
                '<div class="help-tip"><strong>Local edits protected:</strong> Update All pre-scans for items you\\'ve modified and asks whether to keep your edits or overwrite them.</div>' +
                '<h4>Share a Check</h4>' +
                '<p>Click <strong>Share a Check</strong> in the Library header to submit a check from your install for community review. Approved submissions appear in the marketplace for other churches.</p>'
        },
        view_drilldown: {
            title: 'View / Drilldown',
            crumb: 'Help &raquo; View / Drilldown',
            body: '<h4>The View button</h4>' +
                '<p>Appears on any auto/search check with count &gt; 0. Opens an ag-grid table of the actual rows.</p>' +
                '<h4>What you can do</h4>' +
                '<ul><li><strong>Sort</strong> any column by clicking its header</li>' +
                '<li><strong>Filter</strong> with the menu icon on each column header</li>' +
                '<li><strong>Click PeopleId or Name</strong> &mdash; opens that person\\'s TouchPoint profile in a new tab</li>' +
                '<li><strong>Click email</strong> &mdash; opens mail client</li>' +
                '<li><strong>Select rows with checkboxes</strong> &mdash; green action bar appears with bulk actions</li>' +
                '<li><strong>&#127991;&#65039; Tag Selected</strong> &mdash; creates a TouchPoint tag with the selected people. Use the tag to drive emails, follow-up tasks, or further filtering.</li>' +
                '<li><strong>&#8595; CSV</strong> &mdash; export current grid to CSV</li></ul>' +
                '<h4>Auto-detected linking</h4>' +
                '<p>The grid recognizes common columns: PeopleId, Name, OrganizationId, OrganizationName, EmailAddress &mdash; and turns them into clickable links automatically.</p>'
        },
        calendar: {
            title: 'Calendar View',
            crumb: 'Help &raquo; Calendar',
            body: '<h4>What it shows</h4>' +
                '<p>Per-group month grid. Each day shows the checks due that day (or part of an ongoing period).</p>' +
                '<h4>Color coding</h4>' +
                '<ul><li><strong>&#9679; <span style="color:#107c10">Green</span></strong> &mdash; completed in this period</li>' +
                '<li><strong>&#9679; <span style="color:#7a6400">Yellow</span></strong> &mdash; due in current period (not done yet)</li>' +
                '<li><strong>&#9679; <span style="color:#d13438">Red</span></strong> &mdash; skipped (period passed without completion)</li>' +
                '<li><strong>&#9679; <span style="color:#888">Gray</span></strong> &mdash; future / upcoming</li></ul>' +
                '<h4>Click a day</h4>' +
                '<p>Opens a detail panel below the grid showing every check on that date with status badge, who completed it, when, and any notes. For skipped items: "No completion record for this period."</p>' +
                '<p>Use the <strong>Prev / Next</strong> buttons to navigate months. Past view is great for an audit trail of who-did-what.</p>'
        },
        morning_batch: {
            title: 'Morning Batch',
            crumb: 'Help &raquo; Morning Batch',
            body: '<h4>What it does</h4>' +
                '<p>Once per day TouchPoint runs your scheduled batch jobs (early morning, typically). When wired up, it auto-runs all your checks and sends one consolidated reminder email per recipient with everything due that day.</p>' +
                '<h4>Setup</h4>' +
                '<p>Open <strong>Settings</strong> &rarr; <strong>Morning Batch</strong> &rarr; click <strong>Install</strong>. The script adds a managed block to TouchPoint\\'s <code>MorningBatch</code> Special Content that calls into this script. Other entries in MorningBatch are preserved.</p>' +
                '<h4>Pause without uninstalling</h4>' +
                '<p>Toggle <strong>"Send reminder emails when morning batch fires"</strong> off in Settings. The wrapper still runs and pings us each morning (so our "active" indicator stays accurate), but no emails go out. Useful for vacation weeks or temporary holds without touching TouchPoint scheduling.</p>' +
                '<h4>Send Now</h4>' +
                '<p>Force a send right now without waiting for morning batch. Same email content as the scheduled send. Useful for "Did the email actually work?" testing, or to push a one-off if a daily check needs immediate attention.</p>' +
                '<h4>Recipients</h4>' +
                '<p>Per-check email recipients are picked via the person picker in each check. If a check has no specific people, it falls back to the <strong>Default Recipient</strong> (in Settings). Emails are resolved at send time from each person\\'s current TouchPoint profile &mdash; so changing someone\\'s email in TouchPoint auto-propagates.</p>'
        },
        permissions: {
            title: 'Permissions',
            crumb: 'Help &raquo; Permissions',
            body: '<h4>Three levels</h4>' +
                '<ul><li><strong>Admin</strong> (TouchPoint built-in) &mdash; full access including Settings configuration. Always has all permissions regardless of below.</li>' +
                '<li><strong>Edit Roles</strong> &mdash; can add/edit/delete groups and checks, manage the library, install/uninstall morning batch, send reminders manually.</li>' +
                '<li><strong>Complete Roles</strong> &mdash; can mark checks complete, toggle steps, run checks. Read-only otherwise.</li></ul>' +
                '<h4>Configure</h4>' +
                '<p>In <strong>Settings</strong>, set comma-separated TouchPoint role names. Defaults are both <code>Admin</code> only. Examples:</p>' +
                '<ul><li><code>Edit Roles: Admin, Staff, MinistryLeader</code></li>' +
                '<li><code>Complete Roles: Admin, Staff, Volunteer</code></li></ul>' +
                '<div class="help-tip">Editors automatically have Complete permission too &mdash; you don\\'t need to add edit roles to the Complete list.</div>' +
                '<h4>Page access</h4>' +
                '<p>The <code>#roles=Access</code> directive at the top of the script controls who can open the page at all (any logged-in TouchPoint user by default). What they can DO once on the page is then governed by the role config above.</p>'
        }
    };

    function showHelp(topic) {
        var t = HELP_TOPICS[topic] || HELP_TOPICS.toc;
        document.getElementById('hm_title').innerHTML = t.title;
        document.getElementById('hm_breadcrumb').innerHTML = t.crumb || '';
        document.getElementById('hm_body').innerHTML = t.body;
        document.getElementById('hm_footer').style.display = (topic === 'toc') ? 'none' : 'flex';
        document.getElementById('hm').classList.add('on');
    }

    // HTML-safe linkify: handles markdown [label](url) AND bare http(s):// URLs.
    // Always escapes first so user content can never inject HTML.
    function linkify(s) {
        if (s === undefined || s === null) return '';
        var t = esc(s);
        var holders = [];
        function _hold(html) { var i = holders.length; holders.push(html); return '\x01OPSLINK' + i + '\x02'; }
        // Markdown-style links: [label](http://...)
        t = t.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, function(m, label, url) {
            return _hold('<a href="' + url + '" target="_blank" rel="noopener" style="color:#2563eb;text-decoration:underline">' + label + '</a>');
        });
        // Bare URLs (avoid trailing sentence punctuation)
        t = t.replace(/(https?:\/\/[^\s<]+)/g, function(url) {
            var trail = '';
            var m = url.match(/[.,!?:;)\]'"]+$/);
            if (m) { trail = m[0]; url = url.substring(0, url.length - trail.length); }
            return _hold('<a href="' + url + '" target="_blank" rel="noopener" style="color:#2563eb;text-decoration:underline">' + url + '</a>') + trail;
        });
        // Restore placeholders
        t = t.replace(/\x01OPSLINK(\d+)\x02/g, function(_, i) { return holders[parseInt(i, 10)]; });
        return t;
    }
    // Catalog payload helpers — send the browser-fetched marketplace to the
    // server in actions that need it, so server doesn't have to fetch from
    // CF (which gets blocked by Bot Fight Mode on TouchPoint's Azure IPs).
    // --- Person picker (autocomplete TouchPoint people) ---------------
    // Returns an object with setValue(csvPids) / getValue() / clear().
    // opts: containerId, hiddenInputId, multiple (default true), placeholder
    function createPersonPicker(opts) {
        var ct = document.getElementById(opts.containerId);
        if (!ct) return null;
        var multiple = opts.multiple !== false;
        var selected = [];

        ct.innerHTML = '<div class="pp-wrap" id="' + opts.containerId + '_wrap">' +
                       '<input type="text" class="pp-input" id="' + opts.containerId + '_input" placeholder="' + esc(opts.placeholder || 'Search people by name or email...') + '">' +
                       '<div class="pp-dropdown" id="' + opts.containerId + '_dd"></div>' +
                       '</div>';
        var wrap = document.getElementById(opts.containerId + '_wrap');
        var input = document.getElementById(opts.containerId + '_input');
        var dd = document.getElementById(opts.containerId + '_dd');

        function updateHidden() {
            var h = document.getElementById(opts.hiddenInputId);
            if (h) h.value = selected.map(function(p){return p.PeopleId;}).join(',');
        }

        function renderChips() {
            // Remove existing chips, keep input + dd
            var chips = wrap.querySelectorAll('.pp-chip');
            for (var i = 0; i < chips.length; i++) chips[i].parentNode.removeChild(chips[i]);
            selected.forEach(function(p, idx) {
                var chip = document.createElement('span');
                chip.className = 'pp-chip';
                var emHtml = p.Email ? ' <span class="pp-chip-em">' + esc(p.Email) + '</span>' : ' <span class="pp-chip-em" style="color:#d13438">no email</span>';
                chip.innerHTML = esc(p.Name) + emHtml + ' <span class="pp-chip-x">&times;</span>';
                chip.querySelector('.pp-chip-x').onclick = function(e) {
                    e.stopPropagation();
                    selected.splice(idx, 1);
                    renderChips();
                    updateHidden();
                };
                wrap.insertBefore(chip, input);
            });
            input.style.display = (!multiple && selected.length > 0) ? 'none' : '';
        }

        var searchTimer = null;
        var lastQ = '';
        input.addEventListener('input', function() {
            clearTimeout(searchTimer);
            var query = input.value.trim();
            if (query === lastQ) return;
            lastQ = query;
            if (query.length < 2) { dd.innerHTML = ''; dd.classList.remove('on'); return; }
            searchTimer = setTimeout(function() {
                $.post(window.location.pathname, {action:'search_people', q:query}, function(r) {
                    try {
                        var d = JSON.parse(r);
                        if (!d.success) { dd.innerHTML = '<div class="pp-empty">' + esc(d.message || 'Search failed') + '</div>'; dd.classList.add('on'); return; }
                        if (!d.results.length) { dd.innerHTML = '<div class="pp-empty">No matches</div>'; dd.classList.add('on'); return; }
                        var h = '';
                        d.results.forEach(function(p) {
                            h += '<div class="pp-row" data-pid="' + p.PeopleId + '" data-name="' + attr(p.Name) + '" data-email="' + attr(p.Email || '') + '">';
                            h += '<div class="pp-row-nm">' + esc(p.Name) + '</div>';
                            if (p.Email) h += '<div class="pp-row-em">' + esc(p.Email) + '</div>';
                            else h += '<div class="pp-row-noem">No email on file</div>';
                            h += '</div>';
                        });
                        dd.innerHTML = h;
                        dd.classList.add('on');
                        var rows = dd.querySelectorAll('.pp-row');
                        for (var i = 0; i < rows.length; i++) {
                            (function(row) {
                                row.addEventListener('mousedown', function(e) { e.preventDefault(); });
                                row.addEventListener('click', function() {
                                    var p = {
                                        PeopleId: parseInt(row.getAttribute('data-pid'), 10),
                                        Name: row.getAttribute('data-name'),
                                        Email: row.getAttribute('data-email')
                                    };
                                    var dup = false;
                                    for (var j = 0; j < selected.length; j++) {
                                        if (selected[j].PeopleId === p.PeopleId) { dup = true; break; }
                                    }
                                    if (!dup) {
                                        if (!multiple) selected = [];
                                        selected.push(p);
                                        renderChips();
                                        updateHidden();
                                    }
                                    input.value = '';
                                    lastQ = '';
                                    dd.innerHTML = '';
                                    dd.classList.remove('on');
                                    input.focus();
                                });
                            })(rows[i]);
                        }
                    } catch(e) {}
                });
            }, 250);
        });

        input.addEventListener('focus', function() { wrap.classList.add('focus'); });
        input.addEventListener('blur', function() {
            wrap.classList.remove('focus');
            setTimeout(function() { dd.classList.remove('on'); }, 200);
        });

        return {
            setValue: function(csv) {
                selected = [];
                renderChips();
                updateHidden();
                if (!csv) return;
                $.post(window.location.pathname, {action:'resolve_people', pids:String(csv)}, function(r) {
                    try {
                        var d = JSON.parse(r);
                        if (d.success) {
                            selected = d.results || [];
                            renderChips();
                            updateHidden();
                        }
                    } catch(e) {}
                });
            },
            getValue: function() {
                return selected.map(function(p){return p.PeopleId;}).join(',');
            },
            clear: function() {
                selected = [];
                renderChips();
                updateHidden();
            }
        };
    }
    // Lazy-init handles for pickers — created on first use
    var _ckmNotifyPicker = null;
    var _ckmCompleteNotifyPicker = null;
    var _setDefaultPicker = null;

    // Saved-search autocomplete: attaches once to the ckm_search input.
    // As user types, suggest matching saved-search names. Clicking a row
    // fills the input. Doesn't fire for short queries or codes (= < >).
    var _searchAcAttached = false;
    function attachSavedSearchAutocomplete() {
        if (_searchAcAttached) return;
        _searchAcAttached = true;
        var input = document.getElementById('ckm_search');
        var dd = document.getElementById('ckm_search_dd');
        if (!input || !dd) return;
        var timer = null;
        var lastQ = '';
        input.addEventListener('input', function() {
            clearTimeout(timer);
            var q = input.value.trim();
            if (q === lastQ) return;
            lastQ = q;
            // Don't autocomplete obvious condition codes — let user type them
            if (q.length < 2 || /[=<>!]/.test(q)) {
                dd.innerHTML = ''; dd.style.display = 'none'; return;
            }
            timer = setTimeout(function() {
                $.post(window.location.pathname, {action:'search_saved_searches', q:q}, function(r) {
                    try {
                        var d = JSON.parse(r);
                        if (!d.success) return;
                        if (!d.results.length) {
                            dd.innerHTML = '<div class="pp-empty">No saved searches match</div>';
                            dd.style.display = 'block';
                            return;
                        }
                        var h = '';
                        d.results.forEach(function(s) {
                            h += '<div class="pp-row" data-name="' + attr(s.Name) + '">';
                            h += '<div class="pp-row-nm">' + esc(s.Name);
                            if (s.Public) h += ' <span style="background:#e8fde8;color:#107c10;padding:1px 5px;border-radius:6px;font-size:10px;font-weight:600;margin-left:4px">public</span>';
                            h += '</div>';
                            if (s.Owner) h += '<div class="pp-row-em">by ' + esc(s.Owner) + '</div>';
                            h += '</div>';
                        });
                        dd.innerHTML = h;
                        dd.style.display = 'block';
                        var rows = dd.querySelectorAll('.pp-row');
                        for (var i = 0; i < rows.length; i++) {
                            (function(row) {
                                row.addEventListener('mousedown', function(e) { e.preventDefault(); });
                                row.addEventListener('click', function() {
                                    input.value = row.getAttribute('data-name');
                                    lastQ = input.value;
                                    dd.innerHTML = '';
                                    dd.style.display = 'none';
                                    input.focus();
                                });
                            })(rows[i]);
                        }
                    } catch(e) {}
                });
            }, 250);
        });
        input.addEventListener('blur', function() {
            setTimeout(function() { dd.style.display = 'none'; }, 200);
        });
        input.addEventListener('focus', function() {
            // Re-show dropdown if there's existing content matching last query
            if (dd.innerHTML && lastQ === input.value.trim() && !/[=<>!]/.test(input.value)) {
                dd.style.display = 'block';
            }
        });
    }

    function _catalogPayload() {
        return JSON.stringify({items: CATALOG || []});
    }
    function _ensureCatalogLoaded(callback) {
        if (CATALOG && CATALOG.length > 0) { callback(); return; }
        $.get('https://scripts.displaycache.com/api/ops-catalog', function(r) {
            try {
                var data = typeof r === 'string' ? JSON.parse(r) : r;
                CATALOG = data.items || [];
                catalogLoaded = true;
            } catch(e) {}
            callback();
        }).fail(function() { callback(); });
    }

    // Drift detection: true if a catalog-sourced check has been locally edited
    // compared to its mp_snapshot from install/last update.
    function hasLocalEdits(c) {
        var s = c && c.mp_snapshot;
        if (!s || typeof s !== 'object') return false;
        if ((c.sql || '') !== (s.sql || '')) return true;
        if ((c.search || '') !== (s.search || '')) return true;
        if ((c.inst || '') !== (s.inst || '')) return true;
        if ((c.th || 0) !== (s.th || 0)) return true;
        if ((c.cr || 5) !== (s.cr || 5)) return true;
        if ((c.freq || '') !== (s.freq || '')) return true;
        if ((c.cat || '') !== (s.cat || '')) return true;
        if ((c.name || '') !== (s.name || '')) return true;
        if (JSON.stringify(c.steps || []) !== JSON.stringify(s.steps || [])) return true;
        return false;
    }
    var CATALOG = [];
    var catalogLoaded = false;
    var curGroup = null;
    var curFreq = 'all';
    var pendingCid = '';
    var pendingGid = '';
    var ctlgFilter = 'All';
    // Marketplace update tracking — populated by checkMarketplaceUpdates()
    var MP_UPDATE_COUNT = 0;     // installed items with newer rev available
    var MP_NEW_COUNT = 0;        // uninstalled items added to catalog within MP_NEW_WINDOW_DAYS
    var MP_UPDATE_NAMES = [];    // names of items with updates (for banner detail)
    var MP_NEW_WINDOW_DAYS = 14; // an item is "new" for this many days after its approvedAt date
    function isItemNew(item) {
        // True if item's approvedAt is within the last MP_NEW_WINDOW_DAYS days.
        if (!item || !item.approvedAt) return false;
        var t = Date.parse(item.approvedAt);
        if (isNaN(t)) return false;
        var ageDays = (Date.now() - t) / 86400000;
        return ageDays >= 0 && ageDays <= MP_NEW_WINDOW_DAYS;
    }
    var calMonth = '';        // 'YYYY-MM' currently shown in calendar; '' = current month
    var calSelectedDay = null; // 'YYYY-MM-DD' selected day in calendar, or null
    function pad2(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
    function todayKey() {
        var t = new Date();
        return t.getFullYear() + '-' + pad2(t.getMonth() + 1) + '-' + pad2(t.getDate());
    }
    var ICONS = ['&#128202;','&#9961;','&#128176;','&#128100;','&#128273;','&#128197;','&#9881;','&#127968;','&#128231;','&#128640;','&#9889;','&#128295;','&#9733;','&#128203;','&#128220;'];
    var COLORS = ['#0078d4','#107c10','#7a6400','#7c3aed','#d13438','#ea580c','#0891b2','#4f46e5','#be185d','#059669'];
    var CATS = ['People','Involvements','Email','Finance','Facilities','General','Archive'];
    var CAT_ICONS = {'People':'&#128100;','Involvements':'&#128101;','Email':'&#128231;','Finance':'&#128176;','Facilities':'&#127968;','General':'&#9881;','Archive':'&#128451;'};
    var FREQ_LABELS = {'daily':'today','weekly':'this week','monthly':'this month','annual':'this year'};
    function ordSuffix(n) { var s=['th','st','nd','rd']; var v=n%100; return s[(v-20)%10]||s[v]||s[0]; }

    function renderGroups() {
        curGroup = null;
        document.getElementById('v_groups').style.display = 'block';
        document.getElementById('v_detail').style.display = 'none';
        var totalChecks = 0, totalDone = 0, totalIssues = 0, totalGroups = G.length;
        for (var k in GS) { totalChecks += GS[k].total; totalDone += GS[k].done; totalIssues += GS[k].issues; }
        var h = '';
        // App update banner (admin only)
        if (IS_ADMIN && APP_UPDATE_AVAILABLE) {
            h += '<div style="background:#e8f0fd;border:1px solid #cfe3ff;border-radius:8px;padding:10px 14px;margin-bottom:14px;display:flex;align-items:center;gap:10px">';
            h += '<div style="font-size:18px">&#128640;</div>';
            h += '<div style="flex:1;font-size:12px;color:#0078d4">';
            h += '<strong>Operations Checklists update available</strong>';
            h += ' &mdash; you have <code>v' + esc(APP_VERSION) + '</code>, latest is <code>v' + esc(APP_LATEST_VERSION) + '</code>. Your data is preserved.';
            h += '</div>';
            h += '<button id="app_update_btn" class="b bp" onclick="applyAppUpdate()" style="white-space:nowrap">Update Now</button>';
            h += '</div>';
        }
        if (NEEDS_BATCH_SETUP) {
            h += '<div style="background:#fff4ce;border:1px solid #f5c842;border-radius:8px;padding:12px 16px;margin-bottom:14px;display:flex;align-items:flex-start;gap:12px">';
            h += '<div style="font-size:24px;line-height:1">&#9888;</div>';
            h += '<div style="flex:1;font-size:13px;color:#7a6400">';
            if (LAST_BATCH_RUN) {
                h += '<strong>Morning batch may have stopped running.</strong> Last run: ' + esc(LAST_BATCH_RUN) + '. ';
                h += 'Check that MorningBatch is still scheduled in TouchPoint.';
            } else {
                h += '<strong>Morning batch reminders are not active.</strong> ';
                h += 'This script is not in your MorningBatch yet. Open Settings &rarr; Morning Batch and click Install.';
            }
            h += '</div>';
            h += '<button class="b bp" onclick="openSettings()" style="white-space:nowrap">Open Settings</button>';
            h += '</div>';
        } else if (BATCH_PENDING_FIRST_RUN) {
            h += '<div style="background:#e8f0fd;border:1px solid #cfe3ff;border-radius:8px;padding:10px 14px;margin-bottom:14px;display:flex;align-items:center;gap:10px">';
            h += '<div style="font-size:18px">&#8987;</div>';
            h += '<div style="flex:1;font-size:12px;color:#0078d4">';
            h += '<strong>Installed in MorningBatch.</strong> Reminders will start sending after the next morning batch fires.';
            h += '</div>';
            h += '</div>';
        }
        // Marketplace updates banner — admins/editors only
        if (CAN_EDIT && (MP_UPDATE_COUNT > 0 || MP_NEW_COUNT > 0)) {
            h += '<div style="background:#f0f7ff;border:1px solid #cfe3ff;border-radius:8px;padding:10px 14px;margin-bottom:14px;display:flex;align-items:center;gap:10px">';
            h += '<div style="font-size:18px">&#128229;</div>';
            h += '<div style="flex:1;font-size:12px;color:#0078d4">';
            var parts = [];
            if (MP_UPDATE_COUNT > 0) parts.push('<strong>' + MP_UPDATE_COUNT + ' update' + (MP_UPDATE_COUNT > 1 ? 's' : '') + '</strong> available');
            if (MP_NEW_COUNT > 0) parts.push('<strong>' + MP_NEW_COUNT + ' new check' + (MP_NEW_COUNT > 1 ? 's' : '') + '</strong> in marketplace');
            h += parts.join(' &middot; ');
            if (MP_UPDATE_COUNT > 0 && MP_UPDATE_NAMES.length) {
                var preview = MP_UPDATE_NAMES.slice(0, 3).map(esc).join(', ');
                if (MP_UPDATE_NAMES.length > 3) preview += ' (+' + (MP_UPDATE_NAMES.length - 3) + ' more)';
                h += '<div style="font-size:11px;color:#666;margin-top:3px">' + preview + '</div>';
            }
            h += '</div>';
            if (MP_UPDATE_COUNT > 0) h += '<button class="b bp" onclick="updateAllMarketplace(\\'\\')" style="white-space:nowrap">Update All (' + MP_UPDATE_COUNT + ')</button>';
            else h += '<div style="font-size:11px;color:#888">Open a group &rarr; Manage Library</div>';
            h += '</div>';
        }
        h += '<div class="oc-head"><div><h2>Operations Checklists <button class="help-btn" onclick="showHelp(\\'dashboard\\')" title="What is this?">?</button></h2>';
        h += '<div class="oc-sub">' + totalGroups + ' groups &middot; ' + totalChecks + ' checks</div></div>';
        h += '<div style="display:flex;gap:6px">';
        h += '<button class="b bo" onclick="showHelp(\\'toc\\')" title="Help & Tutorials">&#10067; Help</button>';
        if (CAN_EDIT) h += '<button class="b bp" onclick="showGroupModal()">+ New Group</button>';
        h += '<button class="b bo" onclick="openSettings()" title="Settings">&#9881; Settings</button></div></div>';
        h += '<div class="sum">';
        h += '<div class="sm" style="background:#e8fde8"><div class="n" style="color:#107c10">' + totalDone + '</div><div class="l">Completed</div></div>';
        h += '<div class="sm" style="background:#f0f0f0"><div class="n" style="color:#555">' + totalChecks + '</div><div class="l">Total Checks</div></div>';
        h += '<div class="sm" style="background:#fde8e8"><div class="n" style="color:#d13438">' + totalIssues + '</div><div class="l">Issues Found</div></div>';
        h += '<div class="sm" style="background:#e8f0fd"><div class="n" style="color:#0078d4">' + totalGroups + '</div><div class="l">Groups</div></div>';
        h += '</div><div class="grd">';
        G.forEach(function(g) {
            var s = GS[g.id] || {total:0,done:0,issues:0,ok:0,warn:0,crit:0};
            var pct = s.total > 0 ? Math.round(s.done * 100 / s.total) : 0;
            var barClr = pct >= 80 ? '#107c10' : pct >= 50 ? '#7a6400' : '#d13438';
            if (s.total === 0) barClr = '#ccc';
            h += '<div class="gc" onclick="openGroup(\\'' + esc(g.id) + '\\')">';
            h += '<div class="gc-top"><div class="gc-icon" style="background:' + esc(g.color) + '20;color:' + esc(g.color) + '">' + g.icon + '</div>';
            h += '<div class="gc-info"><div class="gc-nm">' + esc(g.name);
            if (g.best) h += '<span class="gc-badge">best practice</span>';
            h += '</div><div class="gc-desc">' + esc(g.desc || '') + '</div></div>';
            h += '<div class="gc-acts" onclick="event.stopPropagation()">';
            if (CAN_EDIT) {
                h += '<button onclick="editGroup(\\'' + esc(g.id) + '\\')">&#9998;</button>';
                if (!g.best) h += '<button class="del" onclick="deleteGroup(\\'' + esc(g.id) + '\\')">&#10005;</button>';
            }
            h += '</div></div>';
            h += '<div class="gc-bar"><div class="gc-fill" style="width:' + pct + '%;background:' + barClr + '"></div></div>';
            h += '<div class="gc-bot"><div class="gc-stat">';
            if (s.ok > 0) h += '<div class="gc-s"><div class="gc-dot" style="background:#107c10"></div>' + s.ok + ' pass</div>';
            if (s.warn > 0) h += '<div class="gc-s"><div class="gc-dot" style="background:#7a6400"></div>' + s.warn + ' warn</div>';
            if (s.crit > 0) h += '<div class="gc-s"><div class="gc-dot" style="background:#d13438"></div>' + s.crit + ' crit</div>';
            if (s.ok === 0 && s.warn === 0 && s.crit === 0) h += '<div class="gc-s" style="color:#ccc">No results yet</div>';
            h += '</div><div>' + s.done + '/' + s.total + ' done</div></div></div>';
        });
        h += '</div>';
        h += '<h3 style="font-size:14px;color:#555;margin:20px 0 10px">Recent Activity</h3>';
        var recent = [];
        for (var k in LD) recent.push(LD[k]);
        recent.sort(function(a, b) { return (b.at || '').localeCompare(a.at || ''); });
        recent = recent.slice(0, 10);
        if (recent.length === 0) {
            h += '<p style="color:#999;font-size:13px">No activity yet. Open a group to start checking off items.</p>';
        } else {
            recent.forEach(function(x) {
                var nm = findCheckName(x.checkId);
                var gn = findGroupName(x.groupId);
                h += '<div class="le"><strong>' + esc(nm) + '</strong>';
                if (gn) h += ' <span style="color:#aaa">(' + esc(gn) + ')</span>';
                h += ' by <strong>' + esc(x.by) + '</strong> <span style="color:#999">' + esc(x.at) + '</span>';
                if (x.notes) h += '<div style="color:#555;margin-top:3px">' + linkify(x.notes) + '</div>';
                h += '</div>';
            });
        }
        document.getElementById('v_groups').innerHTML = h;
    }

    function findCheckName(cid) {
        for (var gid in AC) { var checks = AC[gid]; for (var i = 0; i < checks.length; i++) { if (checks[i].id === cid) return checks[i].name; } }
        return cid;
    }
    function findGroupName(gid) {
        for (var i = 0; i < G.length; i++) { if (G[i].id === gid) return G[i].name; }
        return '';
    }

    function openGroup(gid) {
        curGroup = gid; curFreq = 'all';
        document.getElementById('v_groups').style.display = 'none';
        document.getElementById('v_detail').style.display = 'block';
        renderDetail();
    }

    function renderDetail() {
        var g = null;
        for (var i = 0; i < G.length; i++) { if (G[i].id === curGroup) { g = G[i]; break; } }
        if (!g) { renderGroups(); return; }
        var checks = AC[g.id] || [];
        var s = GS[g.id] || {total:0,done:0,issues:0,ok:0,warn:0,crit:0};
        var h = '';
        h += '<div class="bc"><a onclick="renderGroups()">Groups</a><span class="sep">&#9656;</span><span>' + esc(g.name) + '</span></div>';
        h += '<div class="oc-head"><div><h2 style="display:flex;align-items:center;gap:8px"><span style="font-size:28px">' + g.icon + '</span> ' + esc(g.name) + ' <button class="help-btn" onclick="showHelp(\\'group_detail\\')" title="About this view">?</button></h2>';
        if (g.desc) h += '<div class="oc-sub">' + esc(g.desc) + '</div>';
        if (g.owner) h += '<div class="oc-sub">Owner: ' + esc(g.owner) + '</div>';
        h += '</div><div style="display:flex;gap:6px;flex-wrap:wrap">';
        if (CAN_COMPLETE) h += '<button class="b bp" onclick="runAllInGroup()">&#9654; Run All Checks</button>';
        if (CAN_EDIT) {
            var libBadge = '';
            if (MP_UPDATE_COUNT > 0) libBadge += ' <span style="background:#d13438;color:#fff;padding:1px 6px;border-radius:8px;font-size:10px;font-weight:700;margin-left:2px" title="' + MP_UPDATE_COUNT + ' update' + (MP_UPDATE_COUNT > 1 ? 's' : '') + ' available">' + MP_UPDATE_COUNT + '</span>';
            if (MP_NEW_COUNT > 0) libBadge += ' <span style="background:#0078d4;color:#fff;padding:1px 6px;border-radius:8px;font-size:10px;font-weight:700;margin-left:2px" title="' + MP_NEW_COUNT + ' new community check' + (MP_NEW_COUNT > 1 ? 's' : '') + '">' + MP_NEW_COUNT + ' new</span>';
            var libTitle = [];
            if (MP_UPDATE_COUNT > 0) libTitle.push(MP_UPDATE_COUNT + ' update' + (MP_UPDATE_COUNT > 1 ? 's' : '') + ' available');
            if (MP_NEW_COUNT > 0) libTitle.push(MP_NEW_COUNT + ' new community check' + (MP_NEW_COUNT > 1 ? 's' : ''));
            h += '<button class="b bt" onclick="showCatalog()" title="' + (libTitle.length ? libTitle.join(' · ') : 'Browse the community marketplace') + '">Manage Library' + libBadge + '</button>';
            h += '<button class="b bo" onclick="showCheckModal()">+ Add Check</button>';
            h += '<button class="b bo" onclick="resetGroup()" title="Reset all step progress">&#8634; Reset</button>';
        }
        h += '</div></div>';
        h += '<div class="sum">';
        h += '<div class="sm" style="background:#e8fde8"><div class="n" style="color:#107c10">' + s.ok + '</div><div class="l">Passing</div></div>';
        h += '<div class="sm" style="background:#fff4ce"><div class="n" style="color:#7a6400">' + s.warn + '</div><div class="l">Warning</div></div>';
        h += '<div class="sm" style="background:#fde8e8"><div class="n" style="color:#d13438">' + s.crit + '</div><div class="l">Critical</div></div>';
        h += '<div class="sm" style="background:#f0f0f0"><div class="n" style="color:#555">' + s.done + '/' + s.total + '</div><div class="l">Done</div></div>';
        h += '</div>';
        h += '<div class="freq-tabs">';
        var tabs = ['all','daily','weekly','monthly','annual'];
        for (var t = 0; t < tabs.length; t++) {
            var tb = tabs[t];
            h += '<div class="ft' + (curFreq === tb ? ' a' : '') + '" onclick="setFreq(\\'' + tb + '\\')">' + tb.charAt(0).toUpperCase() + tb.slice(1) + '</div>';
        }
        h += '<div class="ft' + (curFreq === 'calendar' ? ' a' : '') + '" onclick="setFreq(\\'calendar\\')">&#128197; Calendar</div>';
        h += '<div class="ft' + (curFreq === 'log' ? ' a' : '') + '" onclick="setFreq(\\'log\\')">History</div>';
        h += '</div>';
        if (curFreq === 'log') {
            h += renderGroupLog(g.id);
        } else if (curFreq === 'calendar') {
            h += renderCalendar(g.id);
        } else {
            var filtered = checks;
            if (curFreq !== 'all') { filtered = []; for (var i = 0; i < checks.length; i++) { if (checks[i].freq === curFreq) filtered.push(checks[i]); } }
            if (filtered.length === 0) {
                h += '<p style="color:#999;font-size:13px;text-align:center;padding:30px">No checks in this category. Click "+ Add Check" or "Manage Library" to add some.</p>';
            } else {
                h += renderChecks(filtered, g.id);
            }
        }
        document.getElementById('v_detail').innerHTML = h;
    }

    function setFreq(f) {
        if (f !== 'calendar') { calMonth = ''; calSelectedDay = null; }
        curFreq = f;
        renderDetail();
    }

    function renderGroupLog(gid) {
        var entries = [];
        for (var k in LD) { if (LD[k].groupId === gid) entries.push(LD[k]); }
        entries.sort(function(a, b) { return (b.at || '').localeCompare(a.at || ''); });
        var h = '<h3 style="font-size:14px;color:#555;margin-bottom:10px">Completion History</h3>';
        if (entries.length === 0) { h += '<p style="color:#999;font-size:13px">No history yet for this group.</p>'; }
        else { entries.forEach(function(x) {
            var nm = findCheckName(x.checkId);
            h += '<div class="le"><strong>' + esc(nm) + '</strong> by <strong>' + esc(x.by) + '</strong> <span style="color:#999">' + esc(x.at) + '</span>';
            if (x.notes) h += '<div style="color:#555;margin-top:3px">' + linkify(x.notes) + '</div>';
            h += '</div>';
        }); }
        return h;
    }

    // --- Calendar (per-group) ---
    function periodKeyJs(freq, dateStr, dueDay, dueMonth) {
        if (!dateStr) return '';
        var parts = dateStr.split('-');
        var yr = parseInt(parts[0], 10), mo = parseInt(parts[1], 10), d = parseInt(parts[2], 10);
        if (freq === 'daily') return dateStr;
        if (freq === 'weekly') {
            var wd = (dueDay !== null && dueDay !== undefined) ? dueDay : 0;
            var jsDow = new Date(yr, mo - 1, d).getDay();
            var pyDow = (jsDow + 6) % 7;
            var daysSince = (pyDow - wd + 7) % 7;
            var sd = new Date(yr, mo - 1, d - daysSince);
            return sd.getFullYear() + '-' + pad2(sd.getMonth() + 1) + '-' + pad2(sd.getDate());
        }
        if (freq === 'monthly') {
            var dd = (dueDay && dueDay >= 1 && dueDay <= 28) ? dueDay : 1;
            if (d >= dd) return yr + '-' + pad2(mo) + '-' + pad2(dd);
            var pmo = mo - 1, pyr = yr;
            if (pmo < 1) { pmo = 12; pyr--; }
            return pyr + '-' + pad2(pmo) + '-' + pad2(dd);
        }
        if (freq === 'annual') {
            var dm = (dueMonth && dueMonth >= 1 && dueMonth <= 12) ? dueMonth : 1;
            var ddA = (dueDay && dueDay >= 1 && dueDay <= 28) ? dueDay : 1;
            var dueDate = new Date(yr, dm - 1, ddA);
            var thisDate = new Date(yr, mo - 1, d);
            return (thisDate >= dueDate ? yr : yr - 1) + '-' + pad2(dm) + '-' + pad2(ddA);
        }
        return '';
    }

    function computeCalendarItems(gid, yr, mo) {
        var checks = AC[gid] || [];
        var items = {};
        // Index logs by checkId
        var logsByCheck = {};
        LOG.forEach(function(e) {
            if (e.groupId !== gid) return;
            (logsByCheck[e.checkId] = logsByCheck[e.checkId] || []).push(e);
        });
        var dim = new Date(yr, mo, 0).getDate();
        var tk = todayKey();
        checks.forEach(function(c) {
            var dueDates = [];
            if (c.freq === 'daily') {
                // Honor due_dows if set; default = every day
                var allowedDows = null;
                if (c.due_dows) {
                    allowedDows = {};
                    String(c.due_dows).split(',').forEach(function(x) {
                        x = x.trim(); if (x) allowedDows[x] = true;
                    });
                }
                for (var d = 1; d <= dim; d++) {
                    if (allowedDows) {
                        var jsDow = new Date(yr, mo - 1, d).getDay();
                        var pyDow = (jsDow + 6) % 7;
                        if (!allowedDows[String(pyDow)]) continue;
                    }
                    dueDates.push(yr + '-' + pad2(mo) + '-' + pad2(d));
                }
            } else if (c.freq === 'weekly') {
                var dueDow = (c.due_day !== null && c.due_day !== undefined) ? c.due_day : 0;
                for (var d2 = 1; d2 <= dim; d2++) {
                    var jsDow = new Date(yr, mo - 1, d2).getDay();
                    var pyDow = (jsDow + 6) % 7;
                    if (pyDow === dueDow) dueDates.push(yr + '-' + pad2(mo) + '-' + pad2(d2));
                }
            } else if (c.freq === 'monthly') {
                var dd = (c.due_day && c.due_day >= 1) ? c.due_day : 1;
                if (dd <= dim) dueDates.push(yr + '-' + pad2(mo) + '-' + pad2(dd));
            } else if (c.freq === 'annual') {
                var dm = (c.due_month && c.due_month >= 1) ? c.due_month : 1;
                var ddA = (c.due_day && c.due_day >= 1) ? c.due_day : 1;
                if (dm === mo && ddA <= dim) dueDates.push(yr + '-' + pad2(mo) + '-' + pad2(ddA));
            }
            dueDates.forEach(function(dateStr) {
                var pkey = periodKeyJs(c.freq, dateStr, c.due_day, c.due_month);
                var matchingLog = null;
                (logsByCheck[c.id] || []).forEach(function(e) {
                    var entryDate = e.at ? e.at.substring(0, 10) : '';
                    if (!entryDate) return;
                    if (periodKeyJs(c.freq, entryDate, c.due_day, c.due_month) === pkey) {
                        if (!matchingLog || (e.at || '') > (matchingLog.at || '')) matchingLog = e;
                    }
                });
                var status;
                var curPkey = periodKeyJs(c.freq, tk, c.due_day, c.due_month);
                if (matchingLog) status = 'done';
                else if (dateStr > tk) status = 'upcoming';
                else if (pkey === curPkey) status = 'due';
                else status = 'skipped';
                (items[dateStr] = items[dateStr] || []).push({check: c, status: status, log: matchingLog});
            });
        });
        return items;
    }

    function renderCalendar(gid) {
        if (!calMonth) {
            var t = new Date();
            calMonth = t.getFullYear() + '-' + pad2(t.getMonth() + 1);
        }
        var parts = calMonth.split('-');
        var yr = parseInt(parts[0], 10), mo = parseInt(parts[1], 10);
        var firstDay = new Date(yr, mo - 1, 1);
        var dim = new Date(yr, mo, 0).getDate();
        var startCol = (firstDay.getDay() + 6) % 7; // 0=Mon
        var monthName = ['January','February','March','April','May','June','July','August','September','October','November','December'][mo - 1];
        var items = computeCalendarItems(gid, yr, mo);
        var tk = todayKey();

        // Month summary counts
        var sumDone = 0, sumDue = 0, sumSkip = 0, sumUp = 0;
        for (var k in items) items[k].forEach(function(it) {
            if (it.status === 'done') sumDone++;
            else if (it.status === 'due') sumDue++;
            else if (it.status === 'skipped') sumSkip++;
            else sumUp++;
        });

        var h = '';
        h += '<div style="display:flex;justify-content:space-between;align-items:center;margin:8px 0 12px">';
        h += '<button class="b bo" onclick="calNav(-1)">&#8249; Prev</button>';
        h += '<div style="text-align:center"><h3 style="margin:0;font-size:16px">' + monthName + ' ' + yr + '</h3>';
        h += '<div style="font-size:11px;color:#666;margin-top:2px">';
        h += '<span style="color:#107c10">&#9679; ' + sumDone + ' done</span> &middot; ';
        h += '<span style="color:#7a6400">&#9679; ' + sumDue + ' due</span> &middot; ';
        h += '<span style="color:#d13438">&#9679; ' + sumSkip + ' skipped</span> &middot; ';
        h += '<span style="color:#888">&#9679; ' + sumUp + ' upcoming</span>';
        h += '</div></div>';
        h += '<button class="b bo" onclick="calNav(1)">Next &#8250;</button>';
        h += '</div>';

        h += '<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:#e0e0e0;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden">';
        ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].forEach(function(dn) {
            h += '<div style="background:#f8f9fa;padding:5px 6px;font-size:11px;font-weight:700;text-align:center;color:#666">' + dn + '</div>';
        });
        for (var i = 0; i < startCol; i++) h += '<div style="background:#fafafa;min-height:78px"></div>';
        for (var d = 1; d <= dim; d++) {
            var dayKey = yr + '-' + pad2(mo) + '-' + pad2(d);
            var dayItems = items[dayKey] || [];
            var isToday = (dayKey === tk);
            var isSel = (dayKey === calSelectedDay);
            var bg = isSel ? '#e8f0fd' : (isToday ? '#fffae6' : '#fff');
            var border = isToday ? 'border:2px solid #f5c842;' : '';
            h += '<div style="background:' + bg + ';min-height:78px;padding:4px 5px;cursor:pointer;' + border + 'position:relative" onclick="calSelectDay(\\'' + dayKey + '\\')">';
            h += '<div style="font-size:11px;font-weight:600;color:#666">' + d + '</div>';
            // Up to 3 dot+name lines, rest as +N
            var maxShow = 3;
            dayItems.slice(0, maxShow).forEach(function(it) {
                var clr = it.status === 'done' ? '#107c10' : it.status === 'skipped' ? '#d13438' : it.status === 'due' ? '#7a6400' : '#888';
                h += '<div style="font-size:10px;color:' + clr + ';white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.2;margin-top:1px" title="' + attr(it.check.name) + '">&#9679; ' + esc(it.check.name) + '</div>';
            });
            if (dayItems.length > maxShow) h += '<div style="font-size:10px;color:#888;margin-top:1px">+' + (dayItems.length - maxShow) + ' more</div>';
            h += '</div>';
        }
        h += '</div>';

        // Selected day detail panel
        if (calSelectedDay && items[calSelectedDay]) {
            h += '<div style="margin-top:14px;padding:12px;background:#f8f9fa;border-radius:8px;border:1px solid #e0e0e0">';
            h += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">';
            h += '<h4 style="margin:0;font-size:14px;color:#333">' + esc(calSelectedDay) + '</h4>';
            h += '<button class="b bo" style="font-size:11px;padding:3px 8px" onclick="calSelectedDay=null;renderDetail()">Close</button>';
            h += '</div>';
            items[calSelectedDay].forEach(function(it) {
                var sBg, sFg, sLabel;
                if (it.status === 'done') { sBg = '#e8fde8'; sFg = '#107c10'; sLabel = 'completed'; }
                else if (it.status === 'skipped') { sBg = '#fde8e8'; sFg = '#d13438'; sLabel = 'skipped'; }
                else if (it.status === 'due') { sBg = '#fff4ce'; sFg = '#7a6400'; sLabel = 'due'; }
                else { sBg = '#f0f0f0'; sFg = '#888'; sLabel = 'upcoming'; }
                h += '<div style="padding:8px 10px;margin-bottom:5px;background:#fff;border-left:3px solid ' + sFg + ';border-radius:0 6px 6px 0;font-size:12px">';
                h += '<div style="font-weight:600;color:#333">' + esc(it.check.name);
                h += ' <span style="background:' + sBg + ';color:' + sFg + ';padding:1px 7px;border-radius:8px;font-size:10px;margin-left:4px;font-weight:600">' + sLabel + '</span>';
                h += ' <span style="color:#888;font-size:11px;font-weight:400">' + esc(it.check.freq) + ' &middot; ' + esc(it.check.cat || 'General') + '</span>';
                h += '</div>';
                if (it.log) {
                    h += '<div style="color:#666;margin-top:3px">by <strong>' + esc(it.log.by) + '</strong> on ' + esc(it.log.at) + '</div>';
                    if (it.log.notes) h += '<div style="color:#555;margin-top:2px;font-style:italic">&ldquo;' + linkify(it.log.notes) + '&rdquo;</div>';
                } else if (it.status === 'skipped') {
                    h += '<div style="color:#888;margin-top:3px">No completion record for this period.</div>';
                }
                h += '</div>';
            });
            h += '</div>';
        } else if (calSelectedDay) {
            h += '<div style="margin-top:14px;padding:14px;background:#f8f9fa;border-radius:8px;text-align:center;color:#888;font-size:12px">No checks scheduled on ' + esc(calSelectedDay) + '.</div>';
        }
        return h;
    }

    function calNav(delta) {
        if (!calMonth) {
            var t = new Date();
            calMonth = t.getFullYear() + '-' + pad2(t.getMonth() + 1);
        }
        var parts = calMonth.split('-');
        var yr = parseInt(parts[0], 10), mo = parseInt(parts[1], 10) + delta;
        if (mo < 1) { mo = 12; yr--; }
        else if (mo > 12) { mo = 1; yr++; }
        calMonth = yr + '-' + pad2(mo);
        calSelectedDay = null;
        renderDetail();
    }

    function calSelectDay(dayKey) {
        calSelectedDay = (calSelectedDay === dayKey) ? null : dayKey;
        renderDetail();
    }

    function renderChecks(checks, gid) {
        // Group checks by category
        var byCat = {};
        var catOrder = [];
        checks.forEach(function(c) {
            var cat = c.cat || 'General';
            if (!byCat[cat]) { byCat[cat] = []; catOrder.push(cat); }
            byCat[cat].push(c);
        });
        var h = '';
        catOrder.forEach(function(cat) {
            var items = byCat[cat];
            var catIcon = CAT_ICONS[cat] || '&#9881;';
            h += '<div class="cat-hdr">' + catIcon + ' ' + esc(cat) + ' <span class="cat-cnt">(' + items.length + ')</span></div>';
            items.forEach(function(c) {
                h += renderOneCheck(c, gid);
            });
        });
        return h;
    }

    function renderOneCheck(c, gid) {
        var r = R[c.id] || {}, cn = r.count !== undefined ? r.count : -1;
        var lc = LC[c.id];
        var ldAll = LD[c.id];
        var sd = SD[c.id] || {};
        var ic = 'x', ii = '&#9744;';
        var freqLabel = FREQ_LABELS[c.freq] || c.freq;
        var DAYS = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
        var dueLabel = '';
        if (c.freq === 'weekly' && c.due_day !== null && c.due_day !== undefined) dueLabel = ' (due ' + DAYS[c.due_day] + ')';
        else if (c.freq === 'monthly' && c.due_day) dueLabel = ' (due ' + c.due_day + ordSuffix(c.due_day) + ')';
        else if (c.freq === 'annual' && c.due_month) { var MONTHS = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']; dueLabel = ' (due ' + MONTHS[c.due_month] + (c.due_day ? ' ' + c.due_day : '') + ')'; }
        if ((c.type === 'auto' || c.type === 'search') && cn >= 0) {
            if (cn === 0) { ic = 'g'; ii = '&#10003;'; }
            else if (cn >= (c.cr || 5)) { ic = 'r'; ii = '&#9888;'; }
            else if (cn > (c.th || 0)) { ic = 'y'; ii = '&#9888;'; }
        } else if (c.type === 'manual' && lc) { ic = 'g'; ii = '&#10003;'; }
        var st = c.steps || [], ds = 0;
        st.forEach(function(s, i) { if (sd[String(i)] && sd[String(i)].done) ds++; });
        var pr = st.length > 0 ? Math.round(ds * 100 / st.length) : 0;
        var bg = '<span class="bg bf">' + esc(c.freq) + esc(dueLabel) + '</span>';
        if (c.type === 'auto') bg += '<span class="bg ba">SQL</span>';
        if (c.type === 'search') bg += '<span class="bg bc2">search</span>';
        if (c.best) bg += '<span class="bg bb">best practice</span>';
        if (lc) bg += '<span class="bg" style="background:#e8fde8;color:#107c10">done ' + esc(freqLabel) + '</span>';
        else if (ldAll) bg += '<span class="bg" style="background:#fff4ce;color:#7a6400">due</span>';
        var ch = '';
        if (c.type === 'auto' || c.type === 'search') {
            if (cn < 0) ch = '<div class="cnt" style="color:#ccc">--</div>';
            else if (cn === 0) ch = '<div class="cnt z">0</div>';
            else if (cn >= (c.cr || 5)) ch = '<div class="cnt c">' + cn + '</div>';
            else ch = '<div class="cnt w">' + cn + '</div>';
        }
        var mt = '';
        if (r.lastRun) mt += 'Checked: ' + esc(r.lastRun);
        if (lc) mt += (mt ? ' | ' : '') + 'Done ' + esc(freqLabel) + ': ' + esc(lc.at) + ' by ' + esc(lc.by);
        else if (ldAll) mt += (mt ? ' | ' : '') + 'Last done: ' + esc(ldAll.at) + ' by ' + esc(ldAll.by);
        var cidJs = esc(c.id), gidJs = esc(gid);
        var h = '<div class="chk">';
        h += '<div class="rw" onclick="toggleDet(\\'' + cidJs + '\\')">';
        h += '<div class="ic ' + ic + '">' + ii + '</div>';
        h += '<div class="inf"><div class="nm">' + esc(c.name) + bg + '</div>';
        h += '<div class="mt">' + mt + '</div>';
        if (st.length > 0) h += '<div class="prg"><div class="prb" style="width:' + pr + '%;background:' + (pr === 100 ? '#107c10' : '#0078d4') + '"></div></div>';
        h += '</div>' + ch + '<div class="acts">';
        if (CAN_COMPLETE && (c.type === 'auto' || c.type === 'search')) h += '<button class="b bo" onclick="event.stopPropagation();runSingle(\\'' + gidJs + '\\',\\'' + cidJs + '\\')">Run</button>';
        // View button:
        //  - auto checks with count > 0: opens ag-grid (auto-runs if no detail cached)
        //  - search checks (condition code): opens ag-grid (server runs QueryList)
        //  - search checks (saved-search name, no = < > in code): opens TouchPoint's
        //    /Query/<name> page since the saved search has its own UI
        if (c.type === 'auto' && cn > 0) h += '<button class="b bp" data-cid="' + attr(c.id) + '" data-gid="' + attr(gid) + '" data-name="' + attr(c.name) + '" onclick="event.stopPropagation();openDetailViewer(this.getAttribute(\\'data-name\\'),this.getAttribute(\\'data-gid\\'),this.getAttribute(\\'data-cid\\'))">View (' + cn + ')</button>';
        if (c.type === 'search' && c.search && cn > 0) {
            var _isCode = /[=<>!]/.test(c.search);
            if (_isCode) {
                // Only show ag-grid View if QueryList actually populated rows.
                // Some search-builder codes work in QueryCount but error in
                // QueryList — those just get a count badge with no View button.
                if (r.detail && r.detail.length > 0) {
                    h += '<button class="b bp" data-cid="' + attr(c.id) + '" data-gid="' + attr(gid) + '" data-name="' + attr(c.name) + '" onclick="event.stopPropagation();openDetailViewer(this.getAttribute(\\'data-name\\'),this.getAttribute(\\'data-gid\\'),this.getAttribute(\\'data-cid\\'))">View (' + cn + ')</button>';
                }
            } else {
                h += '<a class="b bp" style="text-decoration:none" href="/Query/' + encodeURIComponent(c.search) + '" target="_blank" onclick="event.stopPropagation()">View (' + cn + ')</a>';
            }
        }
        if (CAN_EDIT) h += '<button class="b bo" onclick="event.stopPropagation();editCheck(\\'' + gidJs + '\\',\\'' + cidJs + '\\')">&#9998;</button>';
        if (CAN_COMPLETE) h += '<button class="b bs" data-cid="' + attr(c.id) + '" data-gid="' + attr(gid) + '" data-name="' + attr(c.name) + '" onclick="event.stopPropagation();showCompleteFromBtn(this)">&#10003;</button>';
        if (CAN_EDIT) h += '<button class="b bd" onclick="event.stopPropagation();deleteCheck(\\'' + gidJs + '\\',\\'' + cidJs + '\\')">&#10005;</button>';
        h += '</div></div>';
        h += '<div class="det" id="d_' + cidJs + '">';
        if (c.inst) h += '<div class="ins">' + linkify(c.inst) + '</div>';
        if (st.length > 0) {
            st.forEach(function(s, i) {
                var x = sd[String(i)] || {};
                var disabled = CAN_COMPLETE ? '' : ' disabled';
                var metaTitle = '';
                if (x.note && (x.by || x.at)) metaTitle = 'by ' + (x.by || '') + ', ' + (x.at || '');
                h += '<div class="stp">';
                h += '<input type="checkbox" ' + (x.done ? 'checked' : '') + disabled + ' onchange="toggleStep(\\'' + gidJs + '\\',\\'' + cidJs + '\\',' + i + ',this.checked)">';
                h += '<div class="sl">' + linkify(s) + '</div>';
                if (CAN_COMPLETE) {
                    h += '<input class="snp" placeholder="Add note..." value="' + attr(x.note || '') + '"' + (metaTitle ? ' title="' + attr(metaTitle) + '"' : '') + ' onblur="saveStepNote(\\'' + cidJs + '\\',' + i + ',this.value)">';
                } else if (x.note) {
                    h += '<div class="snp" style="color:#555;background:#fafafa"' + (metaTitle ? ' title="' + attr(metaTitle) + '"' : '') + '>' + linkify(x.note) + '</div>';
                }
                if (x.note && (x.by || x.at)) {
                    h += '<span class="sn-meta">&mdash; ' + esc(x.by || '') + (x.at ? ', ' + esc((x.at || '').substring(5, 16)) : '') + '</span>';
                }
                h += '</div>';
            });
        }
        // Recent completion history (last 5 entries for this check)
        var hist = [];
        for (var hi = 0; hi < LOG.length; hi++) {
            if (LOG[hi].checkId === c.id) hist.push(LOG[hi]);
        }
        hist.sort(function(a, b) { return (b.at || '').localeCompare(a.at || ''); });
        var totalHist = hist.length;
        hist = hist.slice(0, 5);
        if (hist.length > 0) {
            h += '<div class="hist-hdr"><span>&#128340; Recent Completions</span><span class="hist-cnt">showing ' + hist.length + ' of ' + totalHist + '</span></div>';
            hist.forEach(function(e) {
                h += '<div class="hist-row">';
                h += '<span class="hist-by">' + esc(e.by || 'Unknown') + '</span>';
                h += '<span class="hist-at">' + esc(e.at || '') + '</span>';
                if (e.notes) h += '<div class="hist-notes">&ldquo;' + linkify(e.notes) + '&rdquo;</div>';
                h += '</div>';
            });
        }
        h += '</div></div>';
        return h;
    }

    function showCompleteFromBtn(btn) {
        showComplete(btn.getAttribute('data-gid'), btn.getAttribute('data-cid'), btn.getAttribute('data-name'));
    }

    function toggleDet(id) { var e = document.getElementById('d_' + id); if (e) e.classList.toggle('op'); }

    // --- Library (enable/disable catalog checks) ---
    function showCatalog() {
        ctlgFilter = 'All';
        document.getElementById('catm').classList.add('on');
        if (catalogLoaded) {
            renderCatalog();
            return;
        }
        document.getElementById('ctlg_list').innerHTML = '<div style="text-align:center;padding:30px;color:#888">Loading marketplace...</div>';
        $.get('https://scripts.displaycache.com/api/ops-catalog', function(r) {
            try {
                var data = typeof r === 'string' ? JSON.parse(r) : r;
                CATALOG = data.items || [];
                catalogLoaded = true;
                renderCatalog();
            } catch(e) {
                document.getElementById('ctlg_list').innerHTML = '<div style="text-align:center;padding:30px;color:#d13438">Failed to load marketplace. Check your connection.</div>';
            }
        }).fail(function() {
            document.getElementById('ctlg_list').innerHTML = '<div style="text-align:center;padding:30px;color:#d13438">Failed to load marketplace. Check your connection.</div>';
        });
    }

    function renderCatalog() {
        var fh = '<div class="ctlg-fb' + (ctlgFilter === 'All' ? ' active' : '') + '" onclick="ctlgFilter=\\'All\\';renderCatalog()">All</div>';
        CATS.forEach(function(cat) {
            fh += '<div class="ctlg-fb' + (ctlgFilter === cat ? ' active' : '') + '" onclick="ctlgFilter=\\'' + cat + '\\';renderCatalog()">' + cat + '</div>';
        });
        document.getElementById('ctlg_filter').innerHTML = fh;

        // Build set of check names already in this group (with revision tracking)
        var existing = {};
        var curChecks = AC[curGroup] || [];
        curChecks.forEach(function(c) { existing[c.name] = { on: true, mp_rev: c.mp_rev || 0, id: c.id }; });

        // Count how many installed items in this group have updates available
        var groupUpdateCount = 0;
        CATALOG.forEach(function(item) {
            var ex = existing[item.name];
            if (ex && item.rev && (item.rev > (ex.mp_rev || 0))) groupUpdateCount++;
        });
        var bar = document.getElementById('ctlg_update_all_bar');
        if (bar) {
            if (groupUpdateCount > 0) {
                document.getElementById('ctlg_update_all_count').textContent = groupUpdateCount;
                bar.style.display = 'flex';
            } else {
                bar.style.display = 'none';
            }
        }

        var byCat = {};
        CATALOG.forEach(function(item, idx) {
            if (ctlgFilter !== 'All' && item.cat !== ctlgFilter) return;
            var cat = item.cat || 'General';
            if (!byCat[cat]) byCat[cat] = [];
            byCat[cat].push({item: item, idx: idx});
        });

        var h = '';
        CATS.forEach(function(cat) {
            if (!byCat[cat]) return;
            var catIcon = CAT_ICONS[cat] || '&#9881;';
            h += '<div class="ctlg-cat"><h4>' + catIcon + ' ' + esc(cat) + ' (' + byCat[cat].length + ')</h4>';
            byCat[cat].forEach(function(entry) {
                var item = entry.item;
                var ex = existing[item.name];
                var isOn = ex ? true : false;
                var hasUpdate = isOn && item.rev && ex.mp_rev && item.rev > ex.mp_rev;
                h += '<label class="ctlg-item">';
                h += '<label class="tgl"><input type="checkbox" name="ctlg_cb" value="' + entry.idx + '"' + (isOn ? ' checked' : '') + '><span class="slider"></span></label>';
                // Look up the actual installed check object so we can detect local edits
                var installedCheck = null;
                if (ex && ex.id) {
                    var curChecksList = AC[curGroup] || [];
                    for (var ix = 0; ix < curChecksList.length; ix++) {
                        if (curChecksList[ix].id === ex.id) { installedCheck = curChecksList[ix]; break; }
                    }
                }
                var isEdited = installedCheck && hasLocalEdits(installedCheck);
                h += '<div style="flex:1"><div class="ctlg-nm">' + esc(item.name);
                if (isOn) h += ' <span class="bg ba">enabled</span>';
                if (isEdited) h += ' <span class="bg" style="background:#fff4ce;color:#7a6400" title="You have local edits to this check vs. the marketplace version">modified locally</span>';
                if (hasUpdate) h += ' <span class="bg" style="background:#fde8e8;color:#d13438">update available</span>';
                if (!isOn && isItemNew(item)) {
                    var addedTxt = item.approvedAt ? new Date(item.approvedAt).toLocaleDateString() : '';
                    h += ' <span class="bg" style="background:#dff6dd;color:#107c10;border:1px solid #92c89e" title="Added to the marketplace ' + addedTxt + ' - badge clears ' + MP_NEW_WINDOW_DAYS + ' days after the add date">&#10024; new</span>';
                }
                if (item.default) h += ' <span class="bg" style="background:#fef3c7;color:#92400e;border:1px solid #fbbf24" title="Curated default - auto-installed in the &quot;' + attr(item.defaultGroup || '') + '&quot; group during initial setup">&#11088; default</span>';
                h += '</div>';
                h += '<div class="ctlg-desc">' + linkify(item.inst || '') + '</div>';
                h += '<div class="ctlg-badges"><span class="bg bf">' + esc(item.freq) + '</span>';
                if (item.type === 'auto') h += '<span class="bg" style="background:#e8fde8;color:#107c10">SQL</span>';
                if (item.type === 'search') h += '<span class="bg bc2">search</span>';
                if (item.community) h += '<span class="bg" style="background:#f3e8fd;color:#7c3aed">community</span>';
                if (item.submittedBy) h += '<span class="bg" style="background:#f0f0f0;color:#888">by ' + esc(item.submittedBy) + '</span>';
                if (item.steps && item.steps.length > 0) h += '<span class="bg" style="background:#f0f0f0;color:#666">' + item.steps.length + ' steps</span>';
                if (hasUpdate) h += ' <button class="b bd" style="margin-left:6px;font-size:10px;padding:2px 8px" data-cid="' + attr(ex.id) + '" data-name="' + attr(item.name) + '" onclick="event.preventDefault();event.stopPropagation();updateFromMarketplace(this.getAttribute(\\'data-cid\\'),this.getAttribute(\\'data-name\\'))">Update</button>';
                h += '</div></div></label>';
            });
            h += '</div>';
        });
        if (!h) h = '<p style="color:#999;font-size:13px;text-align:center">No checks in this category.</p>';
        document.getElementById('ctlg_list').innerHTML = h;
        updateCatalogCount();
    }

    function updateCatalogCount() {
        var total = document.querySelectorAll('input[name=ctlg_cb]').length;
        var on = document.querySelectorAll('input[name=ctlg_cb]:checked').length;
        document.getElementById('ctlg_count').textContent = on + ' of ' + total + ' enabled';
    }

    document.addEventListener('change', function(e) {
        if (e.target && e.target.name === 'ctlg_cb') updateCatalogCount();
    });

    function updateFromMarketplace(cid, name) {
        // Find the local check to detect edits
        var localCheck = null;
        var checks = AC[curGroup] || [];
        for (var i = 0; i < checks.length; i++) {
            if (checks[i].id === cid) { localCheck = checks[i]; break; }
        }
        var prompt;
        if (localCheck && hasLocalEdits(localCheck)) {
            prompt = '\\u26a0 "' + name + '" has LOCAL EDITS that differ from its original marketplace version.\\n\\nUpdating now will OVERWRITE your changes with the new marketplace version. This cannot be undone.\\n\\nProceed?';
        } else {
            prompt = 'Update "' + name + '" with the latest version from the marketplace?';
        }
        if (!confirm(prompt)) return;
        _ensureCatalogLoaded(function() {
            $.post(window.location.pathname, {
                action:'update_from_marketplace',
                group_id: curGroup,
                check_id: cid,
                mp_name: name,
                catalog_json: _catalogPayload()
            }, function(r) {
                try {
                    var d = JSON.parse(r);
                    if (d.success) { catalogLoaded = false; reloadAndStay(); }
                    else alert(d.message || 'Update failed');
                } catch(e) { alert('Update failed'); }
            });
        });
    }

    function updateAllMarketplace(scopeGid) {
        var scope = scopeGid ? 'this group' : 'all groups';
        // Pre-scan client-side to detect locally edited items in scope
        var editedInScope = [];
        for (var gidScan in AC) {
            if (scopeGid && gidScan !== scopeGid) continue;
            var cks = AC[gidScan] || [];
            cks.forEach(function(c) {
                if (c.mp_snapshot && hasLocalEdits(c)) {
                    editedInScope.push({name: c.name, group: findGroupName(gidScan)});
                }
            });
        }
        var skipEdited = false;
        if (editedInScope.length > 0) {
            // Ask: skip or overwrite locally edited items
            var preview = editedInScope.slice(0, 5).map(function(x) { return ' \\u2022 ' + x.name + (x.group ? ' (' + x.group + ')' : ''); }).join('\\n');
            if (editedInScope.length > 5) preview += '\\n \\u2022 ... and ' + (editedInScope.length - 5) + ' more';
            var msg = editedInScope.length + ' check' + (editedInScope.length === 1 ? ' has' : 's have') + ' local edits in ' + scope + ':\\n' + preview + '\\n\\n';
            msg += 'Click OK to KEEP your local edits and skip these items.\\n';
            msg += 'Click Cancel to OVERWRITE them with the marketplace versions.';
            skipEdited = confirm(msg);
            if (!skipEdited) {
                if (!confirm('Are you sure? Local edits to ' + editedInScope.length + ' check' + (editedInScope.length === 1 ? '' : 's') + ' will be replaced. This cannot be undone.')) return;
            }
        } else {
            if (!confirm('Pull the latest marketplace versions for every installed catalog item in ' + scope + ' that has an update available?')) return;
        }
        var data = {action:'update_all_marketplace', skip_edited: skipEdited ? 'true' : 'false', catalog_json: _catalogPayload()};
        if (scopeGid) data.group_id = scopeGid;
        $.post(window.location.pathname, data, function(r) {
            try {
                var d = JSON.parse(r);
                if (d.success) {
                    var rmsg = 'Updated ' + d.updated + ' check' + (d.updated === 1 ? '' : 's') + '.';
                    if (d.skipped > 0) rmsg += '\\nSkipped ' + d.skipped + ' (locally edited).';
                    if (d.details && d.details.length > 0) {
                        rmsg += '\\n\\nUpdated:\\n' + d.details.slice(0, 10).map(function(x) {
                            return '\\u2022 ' + x.name + ' (' + x.group + ') — rev ' + x.from + ' \\u2192 ' + x.to + (x.hadEdits ? ' [overwrote local edits]' : '');
                        }).join('\\n');
                        if (d.details.length > 10) rmsg += '\\n... and ' + (d.details.length - 10) + ' more';
                    }
                    if (d.skippedDetails && d.skippedDetails.length > 0) {
                        rmsg += '\\n\\nSkipped (you can update individually):\\n' + d.skippedDetails.slice(0, 10).map(function(x) {
                            return '\\u2022 ' + x.name + ' (' + x.group + ')';
                        }).join('\\n');
                    }
                    alert(rmsg);
                    catalogLoaded = false;
                    reloadAndStay();
                } else {
                    alert(d.message || 'Update All failed');
                }
            } catch(e) { alert('Update All failed'); }
        });
    }

    function importFromCatalog() {
        var cbs = document.querySelectorAll('input[name=ctlg_cb]:checked');
        var idxs = [];
        cbs.forEach(function(cb) { idxs.push(cb.value); });
        $.post(window.location.pathname, {
            action: 'sync_library',
            group_id: curGroup,
            enabled_indexes: idxs.join(','),
            catalog_json: _catalogPayload()
        }, function(r) {
            try { var d = JSON.parse(r); if (d.success) reloadAndStay(); else alert(d.message || 'Save failed'); } catch(e) {}
        });
    }

    function lookupSearch() {
        var name = document.getElementById('ckm_search').value.trim();
        if (!name) { alert('Enter a search name to look up'); return; }
        var el = document.getElementById('ckm_search_result');
        el.style.display = 'block';
        el.innerHTML = '<span style="color:#888">Looking up "' + name + '"...</span>';
        $.post(window.location.pathname, {action:'lookup_search', search_name:name}, function(r) {
            try {
                var d = JSON.parse(r);
                if (!d.success) { el.innerHTML = '<span style="color:#d13438">' + (d.message || 'Not found') + '</span>'; return; }
                var h = '<div style="margin-bottom:6px"><strong>' + esc(d.name) + '</strong>';
                if (d.owner) h += ' <span style="color:#888;font-size:11px">by ' + esc(d.owner) + '</span>';
                if (d.public) h += ' <span style="background:#e8fde8;color:#107c10;padding:1px 5px;border-radius:6px;font-size:10px;font-weight:600">public</span>';
                h += '</div>';
                if (d.conditions && d.conditions.length > 0) {
                    h += '<div style="font-size:12px;margin-bottom:6px"><strong>Conditions:</strong></div>';
                    h += '<table style="width:100%;font-size:11px;border-collapse:collapse">';
                    h += '<tr style="background:#e8e8e8"><th style="padding:3px 6px;text-align:left">Field</th><th style="padding:3px 6px;text-align:left">Comparison</th><th style="padding:3px 6px;text-align:left">Value</th><th style="padding:3px 6px;text-align:left">Scope</th></tr>';
                    d.conditions.forEach(function(c) {
                        var scope = '';
                        if (c.program) scope += 'Prog: ' + c.program;
                        if (c.division) scope += (scope ? ', ' : '') + 'Div: ' + c.division;
                        if (c.organization) scope += (scope ? ', ' : '') + 'Org: ' + c.organization;
                        h += '<tr><td style="padding:3px 6px;border-bottom:1px solid #f0f0f0">' + esc(c.field) + '</td>';
                        h += '<td style="padding:3px 6px;border-bottom:1px solid #f0f0f0">' + esc(c.comparison) + '</td>';
                        h += '<td style="padding:3px 6px;border-bottom:1px solid #f0f0f0">' + esc(c.value) + '</td>';
                        h += '<td style="padding:3px 6px;border-bottom:1px solid #f0f0f0;color:#888">' + esc(scope) + '</td></tr>';
                    });
                    h += '</table>';
                } else {
                    h += '<div style="color:#888">No parseable conditions found.</div>';
                }
                if (d.xml) {
                    h += '<details style="margin-top:8px"><summary style="cursor:pointer;font-size:11px;color:#0078d4">View raw XML</summary>';
                    h += '<pre style="font-size:10px;background:#fff;padding:8px;border-radius:4px;overflow-x:auto;margin-top:4px;border:1px solid #e0e0e0;white-space:pre-wrap">' + d.xml.replace(/</g,'&lt;').replace(/>/g,'&gt;') + '</pre></details>';
                }
                el.innerHTML = h;
            } catch(e) { el.innerHTML = '<span style="color:#d13438">Error parsing response</span>'; }
        });
    }

    function checkSearchPortable(val) {
        var warn = document.getElementById('sub_search_warn');
        if (!warn) return;
        // If value has no operators, it's likely a saved search name (not portable)
        if (val && val.length > 2 && val.indexOf('=') === -1 && val.indexOf('>') === -1 && val.indexOf('<') === -1) {
            warn.style.display = 'block';
        } else {
            warn.style.display = 'none';
        }
    }

    function submitToMarketplace() {
        var nm = document.getElementById('sub_name').value.trim();
        if (!nm) { alert('Name is required'); return; }
        var btn = event.target; btn.disabled = true; btn.textContent = 'Submitting...';
        $.post(window.location.pathname, {
            action: 'submit_to_marketplace',
            sub_name: nm,
            sub_cat: document.getElementById('sub_cat').value,
            sub_freq: document.getElementById('sub_freq').value,
            sub_type: document.getElementById('sub_type').value,
            sub_sql: document.getElementById('sub_sql').value,
            sub_search: document.getElementById('sub_search').value,
            sub_th: document.getElementById('sub_th').value || '0',
            sub_cr: document.getElementById('sub_cr').value || '5',
            sub_inst: document.getElementById('sub_inst').value,
            sub_steps: document.getElementById('sub_steps').value
        }, function(r) {
            try {
                var d = JSON.parse(r);
                if (d.success) {
                    alert(d.message || 'Submitted successfully! It will appear after review.');
                    document.getElementById('subm').classList.remove('on');
                } else {
                    alert('Error: ' + (d.message || 'Submission failed'));
                }
            } catch(e) { alert('Submission failed. Please try again.'); }
            btn.disabled = false; btn.textContent = 'Submit for Review';
        });
    }

    // --- AJAX calls ---
    function runAllInGroup() {
        var btn = event.target; btn.disabled = true; btn.innerHTML = 'Running...';
        $.post(window.location.pathname, {action:'run_checks', group_id:curGroup}, function(r) {
            try { var d = JSON.parse(r); if (d.success) location.reload(); } catch(e) {}
            btn.disabled = false; btn.innerHTML = '&#9654; Run All Checks';
        });
    }
    function runSingle(gid, cid) {
        $.post(window.location.pathname, {action:'run_single', group_id:gid, check_id:cid}, function(r) {
            try { var d = JSON.parse(r); if (d.success) { R[cid] = {count:d.count, lastRun:d.lastRun}; updateGroupStats(gid); renderDetail(); } } catch(e) {}
        });
    }
    function toggleStep(gid, cid, idx, chk) {
        var nt = '';
        try { var els = document.querySelectorAll('#d_' + cid + ' .snp'); if (els && els[idx]) nt = els[idx].value; } catch(e) {}
        $.post(window.location.pathname, {action:'toggle_step', check_id:cid, step_idx:idx, checked:chk?'true':'false', note:nt}, function(r) {
            try { var d = JSON.parse(r); if (d.success) {
                if (!SD[cid]) SD[cid] = {};
                SD[cid][String(idx)] = {done:chk, note:nt, by:'You', at:'Now'};
                renderDetail();
                setTimeout(function() { var e = document.getElementById('d_' + cid); if (e) e.classList.add('op'); }, 50);
            }} catch(e) {}
        });
    }
    function saveStepNote(cid, idx, nt) {
        $.post(window.location.pathname, {action:'save_step_note', check_id:cid, step_idx:idx, note:nt}, function(r) {
            try { var d = JSON.parse(r); if (d.success) {
                if (!SD[cid]) SD[cid] = {};
                var k = String(idx);
                if (!SD[cid][k]) SD[cid][k] = {done:false};
                SD[cid][k].note = nt; SD[cid][k].by = 'You'; SD[cid][k].at = 'Now';
            }} catch(e) {}
        });
    }
    function showComplete(gid, cid, nm) {
        pendingCid = cid; pendingGid = gid;
        document.getElementById('cm_name').textContent = nm;
        document.getElementById('cm_notes').value = '';
        document.getElementById('cm').classList.add('on');
    }
    function completeCheck() {
        $.post(window.location.pathname, {action:'complete_check', check_id:pendingCid, group_id:pendingGid, notes:document.getElementById('cm_notes').value}, function(r) {
            try { var d = JSON.parse(r); if (d.success) {
                var entry = {checkId:pendingCid, groupId:pendingGid, by:d.by, at:d.at, notes:document.getElementById('cm_notes').value};
                LD[pendingCid] = entry;
                LC[pendingCid] = entry;
                delete SD[pendingCid];
                document.getElementById('cm').classList.remove('on');
                updateGroupStats(pendingGid); renderDetail();
            }} catch(e) {}
        });
    }
    function resetGroup() {
        if (!confirm('Reset all step progress for this group? Completion history will be kept.')) return;
        $.post(window.location.pathname, {action:'reset_group', group_id:curGroup}, function(r) {
            try { var d = JSON.parse(r); if (d.success) {
                var checks = AC[curGroup] || [];
                checks.forEach(function(c) { delete SD[c.id]; });
                renderDetail();
            }} catch(e) {}
        });
    }

    // --- Group CRUD ---
    function showGroupModal(gid) {
        var g = null;
        if (gid) { for (var i = 0; i < G.length; i++) { if (G[i].id === gid) { g = G[i]; break; } } }
        document.getElementById('gm_title').textContent = g ? 'Edit Group' : 'New Group';
        document.getElementById('gm_id').value = g ? g.id : '';
        document.getElementById('gm_name').value = g ? g.name : '';
        document.getElementById('gm_desc').value = g ? (g.desc || '') : '';
        document.getElementById('gm_owner').value = g ? (g.owner || '') : '';
        var selIcon = g ? g.icon : ICONS[0]; var selClr = g ? g.color : COLORS[0];
        var ih = '';
        ICONS.forEach(function(ic) { ih += '<div class="ic-opt' + (ic === selIcon ? ' sel' : '') + '" onclick="pickIcon(this)" data-icon="' + ic + '">' + ic + '</div>'; });
        document.getElementById('gm_icons').innerHTML = ih;
        var ch = '';
        COLORS.forEach(function(c) { ch += '<div class="clr-opt' + (c === selClr ? ' sel' : '') + '" style="background:' + c + '" onclick="pickColor(this)" data-color="' + c + '"></div>'; });
        document.getElementById('gm_colors').innerHTML = ch;
        document.getElementById('gm').classList.add('on');
    }
    function editGroup(gid) { showGroupModal(gid); }
    function pickIcon(el) { document.querySelectorAll('#gm_icons .ic-opt').forEach(function(x) { x.classList.remove('sel'); }); el.classList.add('sel'); }
    function pickColor(el) { document.querySelectorAll('#gm_colors .clr-opt').forEach(function(x) { x.classList.remove('sel'); }); el.classList.add('sel'); }
    function saveGroup() {
        var nm = document.getElementById('gm_name').value.trim();
        if (!nm) { alert('Name is required'); return; }
        var iconEl = document.querySelector('#gm_icons .ic-opt.sel');
        var clrEl = document.querySelector('#gm_colors .clr-opt.sel');
        $.post(window.location.pathname, {
            action:'save_group', group_id:document.getElementById('gm_id').value,
            group_name:nm, group_desc:document.getElementById('gm_desc').value,
            group_owner:document.getElementById('gm_owner').value,
            group_icon:iconEl ? iconEl.getAttribute('data-icon') : '&#128203;',
            group_color:clrEl ? clrEl.getAttribute('data-color') : '#0078d4'
        }, function(r) { try { var d = JSON.parse(r); if (d.success) location.reload(); } catch(e) {} });
    }
    function deleteGroup(gid) {
        if (!confirm('Delete this group and all its checks?')) return;
        $.post(window.location.pathname, {action:'delete_group', group_id:gid}, function(r) {
            try { var d = JSON.parse(r); if (d.success) location.reload(); } catch(e) {}
        });
    }

    // --- Check CRUD ---
    function showCheckModal() {
        document.getElementById('ckm_title').textContent = 'Add Check';
        document.getElementById('ckm_id').value = '';
        document.getElementById('ckm_gid').value = curGroup;
        document.getElementById('ckm_name').value = '';
        document.getElementById('ckm_cat').value = 'General';
        document.getElementById('ckm_freq').value = 'monthly';
        document.getElementById('ckm_type').value = 'manual';
        document.getElementById('ckm_sql').value = '';
        document.getElementById('ckm_search').value = '';
        document.getElementById('ckm_th').value = '0';
        document.getElementById('ckm_cr').value = '5';
        document.getElementById('ckm_th2').value = '0';
        document.getElementById('ckm_cr2').value = '5';
        document.getElementById('ckm_inst').value = '';
        document.getElementById('ckm_steps').value = '';
        document.getElementById('ckm_sql_wrap').style.display = 'none';
        document.getElementById('ckm_search_wrap').style.display = 'none';
        document.getElementById('ckm_notify_pids').value = '';
        document.getElementById('ckm_complete_notify_pids').value = '';
        if (!_ckmNotifyPicker) _ckmNotifyPicker = createPersonPicker({containerId:'ckm_notify_picker', hiddenInputId:'ckm_notify_pids', placeholder:'Search TouchPoint people...'});
        if (_ckmNotifyPicker) _ckmNotifyPicker.clear();
        if (!_ckmCompleteNotifyPicker) _ckmCompleteNotifyPicker = createPersonPicker({containerId:'ckm_complete_notify_picker', hiddenInputId:'ckm_complete_notify_pids', placeholder:'Search TouchPoint people...'});
        if (_ckmCompleteNotifyPicker) _ckmCompleteNotifyPicker.clear();
        document.getElementById('ckm_due_day_w').value = '0';
        document.getElementById('ckm_due_day_m').value = '1';
        document.getElementById('ckm_due_month_a').value = '1';
        document.getElementById('ckm_due_day_a').value = '1';
        // New daily checks default to weekdays
        setDows('weekdays');
        onFreqChange();
        attachSavedSearchAutocomplete();
        document.getElementById('ckm').classList.add('on');
    }
    function setDows(preset) {
        var sel = preset === 'weekdays' ? [0,1,2,3,4]
                : preset === 'weekends' ? [5,6]
                : preset === 'all'      ? [0,1,2,3,4,5,6]
                : [];
        for (var i = 0; i < 7; i++) {
            var el = document.getElementById('ckm_dow_' + i);
            if (el) el.checked = sel.indexOf(i) >= 0;
        }
    }
    function getDowsCsv() {
        var arr = [];
        for (var i = 0; i < 7; i++) {
            var el = document.getElementById('ckm_dow_' + i);
            if (el && el.checked) arr.push(i);
        }
        return arr.join(',');
    }
    function loadDowsCsv(csv) {
        var set = {};
        if (csv) (csv + '').split(',').forEach(function(x){ x = x.trim(); if (x) set[x] = true; });
        for (var i = 0; i < 7; i++) {
            var el = document.getElementById('ckm_dow_' + i);
            if (el) el.checked = !!set[String(i)];
        }
    }
    function onFreqChange() {
        var f = document.getElementById('ckm_freq').value;
        var wrap = document.getElementById('ckm_due_wrap');
        document.getElementById('ckm_due_daily').style.display = 'none';
        document.getElementById('ckm_due_weekly').style.display = 'none';
        document.getElementById('ckm_due_monthly').style.display = 'none';
        document.getElementById('ckm_due_annual').style.display = 'none';
        if (f === 'daily') { wrap.style.display = 'block'; document.getElementById('ckm_due_daily').style.display = 'block'; }
        else if (f === 'weekly') { wrap.style.display = 'block'; document.getElementById('ckm_due_weekly').style.display = 'block'; }
        else if (f === 'monthly') { wrap.style.display = 'block'; document.getElementById('ckm_due_monthly').style.display = 'block'; }
        else if (f === 'annual') { wrap.style.display = 'block'; document.getElementById('ckm_due_annual').style.display = 'flex'; }
        else { wrap.style.display = 'none'; }
    }
    function onTypeChange() {
        var t = document.getElementById('ckm_type').value;
        document.getElementById('ckm_sql_wrap').style.display = t === 'auto' ? 'block' : 'none';
        document.getElementById('ckm_search_wrap').style.display = t === 'search' ? 'block' : 'none';
    }
    function editCheck(gid, cid) {
        var checks = AC[gid] || [];
        var c = null;
        for (var i = 0; i < checks.length; i++) { if (checks[i].id === cid) { c = checks[i]; break; } }
        if (!c) return;
        document.getElementById('ckm_title').textContent = 'Edit Check';
        document.getElementById('ckm_id').value = c.id;
        document.getElementById('ckm_gid').value = gid;
        document.getElementById('ckm_name').value = c.name || '';
        document.getElementById('ckm_cat').value = c.cat || 'General';
        document.getElementById('ckm_freq').value = c.freq || 'monthly';
        document.getElementById('ckm_type').value = c.type || 'manual';
        document.getElementById('ckm_sql').value = c.sql || '';
        document.getElementById('ckm_search').value = c.search || '';
        document.getElementById('ckm_th').value = c.th !== undefined ? c.th : '0';
        document.getElementById('ckm_cr').value = c.cr !== undefined ? c.cr : '5';
        document.getElementById('ckm_th2').value = c.th !== undefined ? c.th : '0';
        document.getElementById('ckm_cr2').value = c.cr !== undefined ? c.cr : '5';
        document.getElementById('ckm_inst').value = c.inst || '';
        document.getElementById('ckm_steps').value = (c.steps || []).join('\\n');
        if (!_ckmNotifyPicker) _ckmNotifyPicker = createPersonPicker({containerId:'ckm_notify_picker', hiddenInputId:'ckm_notify_pids', placeholder:'Search TouchPoint people...'});
        if (_ckmNotifyPicker) _ckmNotifyPicker.setValue(c.notifyPids || '');
        if (!_ckmCompleteNotifyPicker) _ckmCompleteNotifyPicker = createPersonPicker({containerId:'ckm_complete_notify_picker', hiddenInputId:'ckm_complete_notify_pids', placeholder:'Search TouchPoint people...'});
        if (_ckmCompleteNotifyPicker) _ckmCompleteNotifyPicker.setValue(c.completionNotifyPids || '');
        // Due day/month
        var dd = c.due_day, dm = c.due_month;
        document.getElementById('ckm_due_day_w').value = (c.freq === 'weekly' && dd !== null && dd !== undefined) ? dd : '0';
        document.getElementById('ckm_due_day_m').value = (c.freq === 'monthly' && dd !== null && dd !== undefined) ? dd : '1';
        document.getElementById('ckm_due_month_a').value = (c.freq === 'annual' && dm !== null && dm !== undefined) ? dm : '1';
        document.getElementById('ckm_due_day_a').value = (c.freq === 'annual' && dd !== null && dd !== undefined) ? dd : '1';
        // Days-of-week for daily — empty => every day (legacy), otherwise CSV
        loadDowsCsv(c.due_dows || (c.freq === 'daily' ? '0,1,2,3,4,5,6' : ''));
        onTypeChange();
        onFreqChange();
        attachSavedSearchAutocomplete();
        document.getElementById('ckm').classList.add('on');
    }
    function saveCheck() {
        var nm = document.getElementById('ckm_name').value.trim();
        if (!nm) { alert('Name is required'); return; }
        var tp = document.getElementById('ckm_type').value;
        var freq = document.getElementById('ckm_freq').value;
        var th = tp === 'search' ? document.getElementById('ckm_th2').value : document.getElementById('ckm_th').value;
        var cr = tp === 'search' ? document.getElementById('ckm_cr2').value : document.getElementById('ckm_cr').value;
        // Get due day/month based on frequency
        var dueDay = '', dueMonth = '';
        if (freq === 'weekly') dueDay = document.getElementById('ckm_due_day_w').value;
        else if (freq === 'monthly') dueDay = document.getElementById('ckm_due_day_m').value;
        else if (freq === 'annual') { dueDay = document.getElementById('ckm_due_day_a').value; dueMonth = document.getElementById('ckm_due_month_a').value; }
        $.post(window.location.pathname, {
            action:'save_check', group_id:document.getElementById('ckm_gid').value,
            check_id:document.getElementById('ckm_id').value, check_name:nm,
            check_cat:document.getElementById('ckm_cat').value,
            check_freq:freq, check_type:tp,
            check_sql:document.getElementById('ckm_sql').value,
            check_search:document.getElementById('ckm_search').value,
            check_th:th || '0', check_cr:cr || '5',
            check_due_day:dueDay, check_due_month:dueMonth,
            check_due_dows: (freq === 'daily' ? getDowsCsv() : ''),
            check_notify_pids: document.getElementById('ckm_notify_pids').value,
            check_complete_notify_pids: document.getElementById('ckm_complete_notify_pids').value,
            check_notify: '',
            check_inst:document.getElementById('ckm_inst').value,
            check_steps:document.getElementById('ckm_steps').value
        }, function(r) { try { var d = JSON.parse(r); if (d.success) reloadAndStay(); } catch(e) {} });
    }
    function deleteCheck(gid, cid) {
        if (!confirm('Delete this check?')) return;
        $.post(window.location.pathname, {action:'delete_check', group_id:gid, check_id:cid}, function(r) {
            try { var d = JSON.parse(r); if (d.success) reloadAndStay(); } catch(e) {}
        });
    }

    // --- Helpers ---
    function updateGroupStats(gid) {
        var checks = AC[gid] || [];
        var s = {total:checks.length, done:0, issues:0, ok:0, warn:0, crit:0};
        checks.forEach(function(c) {
            var r = R[c.id] || {}, cn = r.count !== undefined ? r.count : -1;
            var isAuto = (c.type === 'auto' || c.type === 'search');
            if (isAuto && cn >= 0) {
                if (cn === 0) s.ok++;
                else if (cn >= (c.cr || 5)) { s.crit++; s.issues += cn; }
                else if (cn > (c.th || 0)) { s.warn++; s.issues += cn; }
            }
            if (LC[c.id]) s.done++;
            else if (isAuto && cn === 0) s.done++;
        });
        GS[gid] = s;
    }

    // --- Settings ---
    function openSettings() {
        $.post(window.location.pathname, {action:'get_settings'}, function(r) {
            try {
                var d = JSON.parse(r);
                if (d.success) {
                    var s = d.settings || {};
                    document.getElementById('set_email').value = s.defaultEmail || '';
                    if (!_setDefaultPicker) _setDefaultPicker = createPersonPicker({containerId:'set_default_picker', hiddenInputId:'set_default_recipient_pid', multiple:false, placeholder:'Search for default recipient...'});
                    if (_setDefaultPicker) _setDefaultPicker.setValue(s.defaultRecipientPid ? String(s.defaultRecipientPid) : '');
                    document.getElementById('set_cc').checked = s.ccDefault || false;
                    document.getElementById('set_edit_roles').value = s.editRoles || 'Admin';
                    document.getElementById('set_complete_roles').value = s.completeRoles || 'Admin';
                    document.getElementById('set_batch_enabled').checked = (s.batchEnabled !== false);
                    // Batch status indicator
                    var bs = document.getElementById('set_batch_status');
                    var batchOn = (s.batchEnabled !== false);
                    if (bs) {
                        if (NEEDS_BATCH_SETUP) {
                            bs.innerHTML = '<span style="background:#fde8e8;color:#d13438;padding:2px 8px;border-radius:8px;margin-left:6px">not active</span>';
                        } else if (!batchOn) {
                            bs.innerHTML = '<span style="background:#fff4ce;color:#7a6400;padding:2px 8px;border-radius:8px;margin-left:6px">paused</span>' + (s.lastBatchRun ? '<span style="font-size:11px;color:#888;margin-left:6px">last run ' + esc(s.lastBatchRun) + '</span>' : '');
                        } else if (s.lastBatchRun) {
                            bs.innerHTML = '<span style="background:#e8fde8;color:#107c10;padding:2px 8px;border-radius:8px;margin-left:6px">active &middot; last run ' + esc(s.lastBatchRun) + '</span>';
                        } else {
                            bs.innerHTML = '';
                        }
                    }
                    // Inject actual script name into wrapper instructions
                    var sn = SCRIPT_NAME || 'TPxi_OpsChecklists';
                    var wn = document.getElementById('set_wrapper_name');
                    if (wn) wn.textContent = sn + 'Batch';
                    var wc = document.getElementById('set_wrapper_code');
                    if (wc) wc.textContent = 'Data.run_batch = "true"\\nmodel.CallScript("' + sn + '")';
                    var snEl = document.getElementById('set_script_name');
                    if (snEl) snEl.textContent = sn;
                    // Last manual send info
                    var snl = document.getElementById('send_now_last');
                    if (snl) {
                        if (s.lastManualSend) {
                            snl.innerHTML = 'Last manual send: ' + esc(s.lastManualSend) + ' by ' + esc(s.lastManualSendBy || 'Unknown') + (s.lastManualSendResult ? ' &mdash; ' + esc(s.lastManualSendResult) : '');
                        } else {
                            snl.innerHTML = '';
                        }
                    }
                    // Disable inputs and save button for non-admin users
                    var disable = !d.isAdmin;
                    var ids = ['set_email','set_cc','set_edit_roles','set_complete_roles','send_now_btn','set_batch_enabled','install_batch_btn','uninstall_batch_btn'];
                    // Hide the install panel entirely for non-admins
                    var ip = document.getElementById('set_install_panel');
                    if (ip) ip.style.display = disable ? 'none' : '';
                    if (!disable) refreshBatchInstallStatus();
                    ids.forEach(function(id) { var el = document.getElementById(id); if (el) el.disabled = disable; });
                    var saveBtn = document.getElementById('set_save_btn');
                    if (saveBtn) saveBtn.style.display = disable ? 'none' : '';
                    var warn = document.getElementById('set_admin_warn');
                    if (warn) warn.style.display = disable ? 'block' : 'none';
                }
            } catch(e) {}
        });
        document.getElementById('setm').classList.add('on');
    }
    function saveSettings() {
        $.post(window.location.pathname, {
            action:'save_settings',
            default_email:document.getElementById('set_email').value,
            default_recipient_pid: document.getElementById('set_default_recipient_pid').value,
            cc_default:document.getElementById('set_cc').checked ? 'true' : 'false',
            edit_roles:document.getElementById('set_edit_roles').value,
            complete_roles:document.getElementById('set_complete_roles').value,
            batch_enabled:document.getElementById('set_batch_enabled').checked ? 'true' : 'false'
        }, function(r) {
            try {
                var d = JSON.parse(r);
                if (d.success) {
                    document.getElementById('setm').classList.remove('on');
                    location.reload();
                } else {
                    alert(d.message || 'Save failed');
                }
            } catch(e) { alert('Save failed'); }
        });
    }

    // --- SQL detail viewer (ag-grid) -----------------------------------
    var _agLoaded = false;
    var _currentDetailGrid = null;
    var _currentDetailRows = [];
    var _currentDetailCols = [];
    var _currentDetailName = '';
    function loadAgGrid(callback) {
        // Belt-and-suspenders: ensure CSS link tags are in document.head even
        // if TouchPoint stripped the ones we put in model.Form. JS-injected
        // link tags can't be sanitized away.
        function _injectCssIfMissing(href) {
            var existing = document.querySelectorAll('link[rel="stylesheet"]');
            for (var i = 0; i < existing.length; i++) {
                if (existing[i].href === href || existing[i].href.indexOf(href.replace(/^https?:/, '')) >= 0) return false;
            }
            var l = document.createElement('link');
            l.rel = 'stylesheet';
            l.href = href;
            document.head.appendChild(l);
            return true;
        }
        var ag1 = 'https://cdn.jsdelivr.net/npm/ag-grid-community@32.0.2/styles/ag-grid.css';
        var ag2 = 'https://cdn.jsdelivr.net/npm/ag-grid-community@32.0.2/styles/ag-theme-alpine.css';
        var injected1 = _injectCssIfMissing(ag1);
        var injected2 = _injectCssIfMissing(ag2);
        // If CSS was just injected, give it a moment to download/apply.
        var cssDelay = (injected1 || injected2) ? 250 : 0;

        if (_agLoaded || (window.agGrid && window.agGrid.createGrid)) {
            _agLoaded = true;
            setTimeout(callback, cssDelay);
            return;
        }
        var s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/ag-grid-community@32.0.2/dist/ag-grid-community.min.noStyle.js';
        s.onload = function() { _agLoaded = true; setTimeout(callback, cssDelay); };
        s.onerror = function() {
            document.getElementById('dvm_status').innerHTML = '<span style="color:#d13438">ag-grid library failed to load. Falling back to plain table.</span>';
            renderDetailFallback();
        };
        document.head.appendChild(s);
    }
    function openDetailViewer(checkName, gid, cid) {
        _currentDetailName = checkName;
        document.getElementById('dvm_title').textContent = checkName;
        document.getElementById('dvm_meta').innerHTML = 'Loading...';
        document.getElementById('dvm_grid').innerHTML = '<div style="padding:30px;text-align:center;color:#888">Loading detail rows...</div>';
        document.getElementById('dvm_status').innerHTML = '';
        document.getElementById('dvm').classList.add('on');
        _loadDetail(gid, cid, false);
    }

    function _loadDetail(gid, cid, didAutoRun) {
        $.post(window.location.pathname, {action:'get_check_detail', check_id:cid, group_id:gid}, function(r) {
            try {
                var d = JSON.parse(r);
                if (!d.success) { document.getElementById('dvm_grid').innerHTML = '<div style="padding:30px;text-align:center;color:#d13438">Failed to load detail.</div>'; return; }
                _currentDetailRows = d.rows || [];
                _currentDetailCols = d.columns || [];
                // If we have a count > 0 but no rows in cache (e.g. legacy
                // count-only result before the SQL was upgraded), auto-run
                // the check once to populate rows from the new SQL.
                if (!_currentDetailRows.length && !didAutoRun && (typeof d.count === 'number' && d.count > 0)) {
                    document.getElementById('dvm_grid').innerHTML = '<div style="padding:30px;text-align:center;color:#888">Refreshing check to load detail rows...</div>';
                    $.post(window.location.pathname, {action:'run_single', group_id:gid, check_id:cid}, function(r2) {
                        try { JSON.parse(r2); } catch(e) {}
                        _loadDetail(gid, cid, true);
                    }).fail(function() {
                        document.getElementById('dvm_grid').innerHTML = '<div style="padding:30px;text-align:center;color:#d13438">Could not refresh check.</div>';
                    });
                    return;
                }
                var meta = '';
                if (typeof d.count === 'number' && d.count >= 0) meta += d.count + ' total ';
                if (d.lastRun) meta += '&middot; checked ' + esc(d.lastRun) + ' ';
                if (d.truncated) meta += '<span style="color:#7a6400;font-weight:600">&middot; showing first ' + _currentDetailRows.length + '</span>';
                if (d.error) meta = '<span style="color:#d13438">Error: ' + esc(d.error) + '</span>';
                document.getElementById('dvm_meta').innerHTML = meta;
                if (!_currentDetailRows.length) {
                    if (typeof d.count === 'number' && d.count === 0) {
                        document.getElementById('dvm_grid').innerHTML = '<div style="padding:30px;text-align:center;color:#107c10;font-size:14px">&#10003; No rows &mdash; this check is currently passing.</div>';
                    } else {
                        document.getElementById('dvm_grid').innerHTML = '<div style="padding:30px;text-align:center;color:#888">This check is count-only (returns just <code>cnt</code>). Edit the SQL to return rows for grid drilldown, e.g. <code>SELECT TOP 100 PeopleId, Name, ... FROM ...</code></div>';
                    }
                    return;
                }
                loadAgGrid(renderDetailGrid);
            } catch(e) {
                document.getElementById('dvm_grid').innerHTML = '<div style="padding:30px;text-align:center;color:#d13438">Failed to parse response.</div>';
            }
        }).fail(function() {
            document.getElementById('dvm_grid').innerHTML = '<div style="padding:30px;text-align:center;color:#d13438">Network error.</div>';
        });
    }
    function renderDetailGrid() {
        var el = document.getElementById('dvm_grid');
        el.innerHTML = '';
        // Detect a PeopleId-style column so we can link Name and PeopleId to /Person2/
        var pidField = null;
        var pidCandidates = ['PeopleId', 'PersonId', 'peopleid', 'personid'];
        for (var pi = 0; pi < pidCandidates.length; pi++) {
            if (_currentDetailCols.indexOf(pidCandidates[pi]) >= 0) { pidField = pidCandidates[pi]; break; }
        }
        function _personLink(pid, label) {
            var a = document.createElement('a');
            a.href = '/Person2/' + encodeURIComponent(pid);
            a.target = '_blank';
            a.style.cssText = 'color:#2563eb;font-weight:600;text-decoration:none';
            a.textContent = String(label);
            return a;
        }
        function _orgLink(oid, label) {
            var a = document.createElement('a');
            a.href = '/Organization/' + encodeURIComponent(oid);
            a.target = '_blank';
            a.style.cssText = 'color:#2563eb;font-weight:600;text-decoration:none';
            a.textContent = String(label);
            return a;
        }
        var colDefs = _currentDetailCols.map(function(c) {
            var def = {field: c, headerName: c.replace(/([a-z])([A-Z])/g, '$1 $2'), sortable:true, filter:true, resizable:true};
            var cl = c.toLowerCase();
            // PeopleId column → link to person profile
            if (cl === 'peopleid' || cl === 'personid') {
                def.cellRenderer = function(p) {
                    if (!p.value) return '';
                    return _personLink(p.value, p.value);
                };
            }
            // Name column with a PeopleId in same row → link Name to person
            else if ((cl === 'name' || cl === 'name2' || cl === 'donorname') && pidField) {
                def.cellRenderer = function(p) {
                    if (!p.value) return '';
                    var pid = p.data && p.data[pidField];
                    if (!pid) return String(p.value);
                    return _personLink(pid, p.value);
                };
            }
            // OrganizationId column → link to org page
            else if (cl === 'organizationid' || cl === 'orgid') {
                def.cellRenderer = function(p) {
                    if (!p.value) return '';
                    return _orgLink(p.value, p.value);
                };
            }
            // OrganizationName with OrganizationId in same row → link to org
            else if (cl === 'organizationname' && _currentDetailCols.indexOf('OrganizationId') >= 0) {
                def.cellRenderer = function(p) {
                    if (!p.value) return '';
                    var oid = p.data && p.data.OrganizationId;
                    if (!oid) return String(p.value);
                    return _orgLink(oid, p.value);
                };
            }
            // EmailAddress → mailto link
            else if (cl === 'emailaddress' || cl === 'email') {
                def.cellRenderer = function(p) {
                    if (!p.value) return '';
                    var a = document.createElement('a');
                    a.href = 'mailto:' + p.value;
                    a.style.cssText = 'color:#2563eb;text-decoration:none';
                    a.textContent = String(p.value);
                    return a;
                };
            }
            // Numeric column detection (after special-casing — don't right-align IDs we link)
            else if (_currentDetailRows.length) {
                var sample = _currentDetailRows[0][c];
                if (sample !== '' && sample !== null && sample !== undefined && !isNaN(parseFloat(sample))) {
                    def.type = 'numericColumn';
                }
            }
            return def;
        });
        var rows = _currentDetailRows.map(function(r) {
            var out = {};
            _currentDetailCols.forEach(function(c) {
                var v = r[c];
                if (v === '' || v === null || v === undefined) { out[c] = ''; return; }
                var n = parseFloat(v);
                out[c] = (!isNaN(n) && String(n) === String(v)) ? n : v;
            });
            return out;
        });
        if (_currentDetailGrid) { try { _currentDetailGrid.destroy(); } catch(e) {} }
        // If we have a PeopleId column, enable row selection + checkbox in 1st col
        var hasPid = pidField !== null;
        if (hasPid && colDefs.length > 0) {
            colDefs[0].checkboxSelection = true;
            colDefs[0].headerCheckboxSelection = true;
            colDefs[0].headerCheckboxSelectionFilteredOnly = true;
            colDefs[0].minWidth = 130;  // give room for checkbox + value
        }
        // Reset selection UI on each open
        clearGridSelection();
        var gridOptions = {
            columnDefs: colDefs,
            rowData: rows,
            defaultColDef: {sortable:true, filter:true, resizable:true, minWidth:80},
            animateRows: true,
            pagination: true,
            paginationPageSize: 50,
            paginationPageSizeSelector: [25,50,100,500],
            domLayout: 'normal',
            autoSizeStrategy: {type:'fitGridWidth'},
            rowSelection: hasPid ? 'multiple' : undefined,
            suppressRowClickSelection: hasPid,
            onSelectionChanged: hasPid ? function(event) {
                var sel = event.api.getSelectedRows();
                window._dvmSelectedRows = sel;
                window._dvmSelectedPidField = pidField;
                var bar = document.getElementById('dvm_action_bar');
                var cnt = document.getElementById('dvm_sel_count');
                if (bar) bar.style.display = sel.length > 0 ? 'flex' : 'none';
                if (cnt) cnt.textContent = sel.length + ' selected';
            } : undefined
        };
        try {
            _currentDetailGrid = window.agGrid.createGrid(el, gridOptions);
        } catch(e) {
            document.getElementById('dvm_status').innerHTML = '<span style="color:#d13438">ag-grid error: ' + esc(String(e)) + '. Falling back to plain table.</span>';
            renderDetailFallback();
            return;
        }
        // Sanity check — if rows still not rendered after 800ms, fall back
        setTimeout(function() {
            var rendered = el.querySelectorAll('.ag-row, [role="row"]');
            if (!rendered || rendered.length === 0) {
                document.getElementById('dvm_status').innerHTML = '<span style="color:#7a6400">ag-grid rendered empty &mdash; using plain-table fallback.</span>';
                renderDetailFallback();
            }
        }, 800);
    }
    function renderDetailFallback() {
        var el = document.getElementById('dvm_grid');
        if (!_currentDetailRows.length) { el.innerHTML = '<div style="padding:20px;text-align:center;color:#888">No rows.</div>'; return; }
        // Find PeopleId column for linking Name fields
        var pidField = null;
        var pidCandidates = ['PeopleId', 'PersonId', 'peopleid', 'personid'];
        for (var pi = 0; pi < pidCandidates.length; pi++) {
            if (_currentDetailCols.indexOf(pidCandidates[pi]) >= 0) { pidField = pidCandidates[pi]; break; }
        }
        function _cell(c, r) {
            var v = r[c];
            if (v === null || v === undefined || v === '') return '';
            var cl = c.toLowerCase();
            if (cl === 'peopleid' || cl === 'personid') {
                return '<a href="/Person2/' + encodeURIComponent(v) + '" target="_blank" style="color:#2563eb;font-weight:600;text-decoration:none">' + esc(v) + '</a>';
            }
            if ((cl === 'name' || cl === 'name2' || cl === 'donorname') && pidField && r[pidField]) {
                return '<a href="/Person2/' + encodeURIComponent(r[pidField]) + '" target="_blank" style="color:#2563eb;font-weight:600;text-decoration:none">' + esc(v) + '</a>';
            }
            if (cl === 'organizationid' || cl === 'orgid') {
                return '<a href="/Organization/' + encodeURIComponent(v) + '" target="_blank" style="color:#2563eb;font-weight:600;text-decoration:none">' + esc(v) + '</a>';
            }
            if (cl === 'organizationname' && r.OrganizationId) {
                return '<a href="/Organization/' + encodeURIComponent(r.OrganizationId) + '" target="_blank" style="color:#2563eb;font-weight:600;text-decoration:none">' + esc(v) + '</a>';
            }
            if (cl === 'emailaddress' || cl === 'email') {
                return '<a href="mailto:' + esc(v) + '" style="color:#2563eb;text-decoration:none">' + esc(v) + '</a>';
            }
            return esc(v);
        }
        var h = '<div style="overflow:auto;max-height:60vh"><table style="width:100%;border-collapse:collapse;font-size:12px"><thead><tr>';
        _currentDetailCols.forEach(function(c) { h += '<th style="text-align:left;padding:6px 8px;background:#f0f0f0;border-bottom:2px solid #ccc;position:sticky;top:0">' + esc(c) + '</th>'; });
        h += '</tr></thead><tbody>';
        _currentDetailRows.forEach(function(r) {
            h += '<tr>';
            _currentDetailCols.forEach(function(c) { h += '<td style="padding:5px 8px;border-bottom:1px solid #eee">' + _cell(c, r) + '</td>'; });
            h += '</tr>';
        });
        h += '</tbody></table></div>';
        el.innerHTML = h;
    }
    function clearGridSelection() {
        try {
            if (_currentDetailGrid && _currentDetailGrid.deselectAll) _currentDetailGrid.deselectAll();
        } catch(e) {}
        window._dvmSelectedRows = [];
        var bar = document.getElementById('dvm_action_bar');
        if (bar) bar.style.display = 'none';
        var cnt = document.getElementById('dvm_sel_count');
        if (cnt) cnt.textContent = '0 selected';
    }

    function openTagModal() {
        var sel = window._dvmSelectedRows || [];
        if (!sel.length) { alert('No rows selected'); return; }
        document.getElementById('tagm_count').textContent = 'Tagging ' + sel.length + ' person' + (sel.length === 1 ? '' : 's') + ' selected from this grid.';
        document.getElementById('tagm_name').value = '';
        document.getElementById('tagm').classList.add('on');
        setTimeout(function() {
            var i = document.getElementById('tagm_name'); if (i) i.focus();
        }, 50);
    }

    function submitBulkTag() {
        var sel = window._dvmSelectedRows || [];
        var name = document.getElementById('tagm_name').value.trim();
        if (!name) { alert('Tag name required'); return; }
        if (!sel.length) { alert('No rows selected'); return; }
        var pidField = window._dvmSelectedPidField;
        if (!pidField) {
            // Fallback: infer from first row
            var keys = Object.keys(sel[0] || {});
            for (var i = 0; i < keys.length; i++) {
                var lk = keys[i].toLowerCase();
                if (lk === 'peopleid' || lk === 'personid') { pidField = keys[i]; break; }
            }
        }
        if (!pidField) { alert('No PeopleId column found'); return; }
        var pids = [];
        for (var j = 0; j < sel.length; j++) {
            var v = sel[j][pidField];
            if (v) pids.push(String(v));
        }
        if (!pids.length) { alert('No valid PeopleIds in selection'); return; }
        var btn = document.getElementById('tagm_submit_btn');
        if (btn) { btn.disabled = true; btn.textContent = 'Tagging...'; }
        $.post(window.location.pathname, {
            action: 'bulk_tag',
            people_ids: pids.join(','),
            tag_name: name
        }, function(r) {
            try {
                var d = JSON.parse(r);
                if (d.success) {
                    alert(d.message || ('Tagged ' + d.count));
                    document.getElementById('tagm').classList.remove('on');
                    clearGridSelection();
                } else {
                    alert(d.message || 'Tag failed');
                }
            } catch(e) { alert('Tag failed (could not parse response)'); }
            if (btn) { btn.disabled = false; btn.textContent = 'Tag'; }
        }).fail(function() {
            alert('Tag failed (network error)');
            if (btn) { btn.disabled = false; btn.textContent = 'Tag'; }
        });
    }

    function exportGridCsv() {
        if (!_currentDetailRows.length) return;
        var lines = [_currentDetailCols.map(function(c){return '"'+String(c).replace(/"/g,'""')+'"';}).join(',')];
        _currentDetailRows.forEach(function(r) {
            lines.push(_currentDetailCols.map(function(c){
                var v = r[c]; if (v === null || v === undefined) v = '';
                return '"' + String(v).replace(/"/g,'""') + '"';
            }).join(','));
        });
        var csv = lines.join('\\n');
        var blob = new Blob([csv], {type:'text/csv'});
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = (_currentDetailName || 'check') + '_' + new Date().toISOString().substring(0,10) + '.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // --- App version / auto-update -------------------------------------
    var APP_UPDATE_AVAILABLE = false;
    var APP_LATEST_VERSION = '';
    function checkForAppUpdate() {
        // Browser-side fetch hits scripts.displaycache.com (which CF lets through
        // for browsers); server-side update fetch uses workers.dev to avoid CF.
        try {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', DC_API_BASE + '/script-versions', true);
            xhr.timeout = 5000;
            xhr.onreadystatechange = function() {
                if (xhr.readyState !== 4) return;
                if (xhr.status !== 200) return;
                try {
                    var versions = JSON.parse(xhr.responseText);
                    var latest = versions[DC_SCRIPT_ID];
                    if (latest && latest !== APP_VERSION) {
                        APP_UPDATE_AVAILABLE = true;
                        APP_LATEST_VERSION = latest;
                        // If we're on the dashboard view, re-render to show banner
                        if (!curGroup) renderGroups();
                    }
                } catch(e) {}
            };
            xhr.send();
        } catch(e) {}
    }

    function applyAppUpdate() {
        if (!confirm('Update Operations Checklists from v' + APP_VERSION + ' to v' + APP_LATEST_VERSION + '?\\n\\nYour groups, checks, completion history, and settings are all stored separately and will be preserved.')) return;
        var btn = document.getElementById('app_update_btn');
        if (btn) { btn.disabled = true; btn.innerHTML = 'Updating...'; }
        $.post(window.location.pathname, {action:'apply_update'}, function(r) {
            try {
                var d = JSON.parse(r);
                if (d.success) {
                    alert(d.message || 'Updated! Reloading...');
                    window.location.reload(true);
                } else {
                    alert('Update failed: ' + (d.message || 'Unknown error'));
                    if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
                }
            } catch(e) {
                alert('Update failed (could not parse response)');
                if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
            }
        }).fail(function() {
            alert('Update failed (network error)');
            if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
        });
    }

    function checkMarketplaceUpdates() {
        // Fetch marketplace catalog and compare to installed items.
        // Counts items with newer rev (updates) and items not installed (new).
        // Also populates global CATALOG so server-side actions can be sent
        // catalog data via POST (avoids CF bot challenges on TouchPoint's
        // server-side fetches).
        $.get('https://scripts.displaycache.com/api/ops-catalog', function(r) {
            try {
                var data = typeof r === 'string' ? JSON.parse(r) : r;
                var items = data.items || [];
                CATALOG = items;
                catalogLoaded = true;
                // Build map of installed catalog items: name -> mp_rev
                var installed = {};
                for (var gid in AC) {
                    (AC[gid] || []).forEach(function(c) {
                        if (c.src === 'catalog' || c.mp_rev !== undefined) {
                            // Track the *highest* mp_rev installed across groups
                            if (!(c.name in installed) || (c.mp_rev || 0) > installed[c.name]) {
                                installed[c.name] = c.mp_rev || 0;
                            }
                        }
                    });
                }
                var updates = 0, newItems = 0;
                var updateNames = [];
                items.forEach(function(it) {
                    var nm = it.name || '';
                    var rev = it.rev || 1;
                    if (nm in installed) {
                        if (rev > installed[nm]) {
                            updates++;
                            updateNames.push(nm);
                        }
                    } else if (isItemNew(it)) {
                        // Only count uninstalled items added within the new-window
                        // as "new" — older uninstalled items are just uninstalled.
                        newItems++;
                    }
                });
                MP_UPDATE_COUNT = updates;
                MP_NEW_COUNT = newItems;
                MP_UPDATE_NAMES = updateNames;
                // Re-render the active view if counts changed from zero
                if (updates > 0 || newItems > 0) {
                    if (curGroup) renderDetail();
                    else renderGroups();
                }
            } catch(e) {}
        });
    }

    function refreshBatchInstallStatus() {
        var st = document.getElementById('install_batch_status');
        var inst = document.getElementById('install_batch_btn');
        var rem = document.getElementById('uninstall_batch_btn');
        if (st) st.textContent = 'Checking MorningBatch...';
        $.post(window.location.pathname, {action:'check_batch_install'}, function(r) {
            try {
                var d = JSON.parse(r);
                if (!d.success) { if (st) st.textContent = ''; return; }
                if (d.installed) {
                    if (st) st.innerHTML = '<span style="color:#107c10;font-weight:600">&#10003; Installed</span> &mdash; <code>MorningBatch</code> calls <code>' + esc(d.scriptName) + '</code> via the managed block.';
                    if (inst) inst.style.display = 'none';
                    if (rem) rem.style.display = '';
                } else {
                    var msg = '<span style="color:#7a6400;font-weight:600">Not installed</span> &mdash; click Install to add the managed block to <code>MorningBatch</code>.';
                    if (d.referencedOutsideBlock) {
                        msg += '<br><span style="color:#d13438">Note:</span> <code>' + esc(d.scriptName) + '</code> is already referenced elsewhere in MorningBatch. Installing will create a duplicate call &mdash; consider removing the existing reference first.';
                    }
                    if (st) st.innerHTML = msg;
                    if (inst) inst.style.display = '';
                    if (rem) rem.style.display = 'none';
                }
            } catch(e) { if (st) st.textContent = ''; }
        });
    }

    function _showInstallResult(success, message) {
        var res = document.getElementById('install_batch_result');
        if (!res) return;
        res.style.display = 'block';
        if (success) {
            res.style.background = '#e8fde8';
            res.style.color = '#107c10';
        } else {
            res.style.background = '#fde8e8';
            res.style.color = '#d13438';
        }
        res.textContent = message;
    }

    function installBatch() {
        if (!confirm('Add a managed block to TouchPoint\\'s MorningBatch script that calls this script each morning? Your existing MorningBatch content will be preserved.')) return;
        var btn = document.getElementById('install_batch_btn');
        if (btn) { btn.disabled = true; btn.textContent = 'Installing...'; }
        $.post(window.location.pathname, {action:'install_batch'}, function(r) {
            try {
                var d = JSON.parse(r);
                _showInstallResult(d.success, d.message || (d.success ? 'Installed.' : 'Install failed.'));
                if (d.success) refreshBatchInstallStatus();
            } catch(e) {
                _showInstallResult(false, 'Install failed (could not parse response).');
            }
            if (btn) { btn.disabled = false; btn.innerHTML = '&#10004; Install'; }
        }).fail(function() {
            _showInstallResult(false, 'Install failed (network error).');
            if (btn) { btn.disabled = false; btn.innerHTML = '&#10004; Install'; }
        });
    }

    function uninstallBatch() {
        if (!confirm('Remove this script\\'s managed block from MorningBatch? Other entries in MorningBatch will be preserved.')) return;
        var btn = document.getElementById('uninstall_batch_btn');
        if (btn) { btn.disabled = true; btn.textContent = 'Removing...'; }
        $.post(window.location.pathname, {action:'uninstall_batch'}, function(r) {
            try {
                var d = JSON.parse(r);
                _showInstallResult(d.success, d.message || (d.success ? 'Removed.' : 'Remove failed.'));
                if (d.success) refreshBatchInstallStatus();
            } catch(e) {
                _showInstallResult(false, 'Remove failed (could not parse response).');
            }
            if (btn) { btn.disabled = false; btn.innerHTML = '&#10005; Remove'; }
        }).fail(function() {
            _showInstallResult(false, 'Remove failed (network error).');
            if (btn) { btn.disabled = false; btn.innerHTML = '&#10005; Remove'; }
        });
    }

    function sendRemindersNow() {
        if (!confirm('Send reminder emails now for everything due today? This sends real emails immediately.')) return;
        var btn = document.getElementById('send_now_btn');
        var res = document.getElementById('send_now_result');
        if (btn) { btn.disabled = true; btn.textContent = 'Sending...'; }
        if (res) {
            res.style.display = 'block';
            res.style.background = '#f0f0f0';
            res.style.color = '#666';
            res.innerHTML = 'Sending emails...';
        }
        $.post(window.location.pathname, {action:'send_reminders_now'}, function(r) {
            try {
                var d = JSON.parse(r);
                if (d.success) {
                    if (res) {
                        if (d.errors > 0) {
                            res.style.background = '#fff4ce';
                            res.style.color = '#7a6400';
                        } else {
                            res.style.background = '#e8fde8';
                            res.style.color = '#107c10';
                        }
                        var msg = '<div>' + esc(d.message || ('Sent ' + d.sent)) + '</div>';
                        if (d.log && d.log.length > 0) {
                            msg += '<details style="margin-top:6px"><summary style="cursor:pointer;font-size:11px">Show details (' + d.log.length + ')</summary>';
                            msg += '<ul style="margin:4px 0 0 0;padding-left:18px;font-size:11px;color:#555">';
                            d.log.forEach(function(ln) { msg += '<li>' + esc(ln) + '</li>'; });
                            msg += '</ul></details>';
                        }
                        res.innerHTML = msg;
                    }
                } else {
                    if (res) {
                        res.style.background = '#fde8e8';
                        res.style.color = '#d13438';
                        res.innerHTML = esc(d.message || 'Send failed');
                    }
                }
            } catch(e) {
                if (res) {
                    res.style.background = '#fde8e8';
                    res.style.color = '#d13438';
                    res.innerHTML = 'Send failed (could not parse response)';
                }
            }
            if (btn) { btn.disabled = false; btn.innerHTML = '&#9993; Send Now'; }
        }).fail(function() {
            if (res) {
                res.style.background = '#fde8e8';
                res.style.color = '#d13438';
                res.innerHTML = 'Send failed (network error)';
            }
            if (btn) { btn.disabled = false; btn.innerHTML = '&#9993; Send Now'; }
        });
    }

    function reloadAndStay() {
        if (curGroup) {
            window.location.hash = 'group=' + curGroup;
        }
        location.reload();
    }

    // On init: check hash for group to auto-open
    var hashMatch = window.location.hash.match(/group=([^&]+)/);
    if (hashMatch) {
        var gid = hashMatch[1];
        // Verify group exists
        for (var i = 0; i < G.length; i++) {
            if (G[i].id === gid) { openGroup(gid); break; }
        }
        if (!curGroup) renderGroups();
    } else {
        renderGroups();
    }
    // Check marketplace for updates/new items in the background (editors only)
    if (CAN_EDIT) {
        setTimeout(checkMarketplaceUpdates, 200);
    }
    // Check for new version of the script itself (admin only)
    if (IS_ADMIN) {
        setTimeout(checkForAppUpdate, 400);
    }
    """

        model.Script = js
