### TPxi_Dashboards
### A dashboard builder with the power of  Enterprise Reporting.  
### pick a dashboard, or start from a library template,
### arrange tiles on a drag-and-resize grid, split across tabs, and control who
### can see each dashboard.
###
#--------------------------------------------------------------------
# TPxi Operations Checklists v1.1.5
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
###
### --Upload Instructions Start--
### To upload code to Touchpoint, use the following steps.
### 1. Click Admin ~ Advanced ~ Special Content ~ Python
### 2. Click New Python Script File
### 3. Name the file "TPxi_Dashboards" and paste all this code
### 4. Test run to confirm it loads, then add to a menu if desired
### --Upload Instructions End--

import datetime
import json
import re
import time
import traceback

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

# Name of the Enterprise Reporting deployment, as it appears in the URL after
# /PyScriptForm/. Leave blank to look for it: not every church runs it, and
# those that do may have installed it under another name. Report tiles are one
# optional source; TouchPoint widget tiles and built tiles need nothing else.
REPORTS_SCRIPT_DEFAULT = "EnterpriseReporting"
CONTENT_SETTINGS = "TPxi_Dashboards_Settings"
# Scratch space for a chunked script upload. A whole script in one POST can
# exceed the request limit once URL-encoded, and ASP.NET rejects a body that
# looks like markup, so it arrives entity-encoded and in pieces.
UPLOAD_KEY = "TPxi_Dashboards_Upload"
# The report catalog, cached locally. Tiles run from THIS copy, never from a
# live fetch: a dashboard that stops working because a CDN is slow is worse
# than one showing a catalog that is a day old.
CONTENT_REPORTS = "TPxi_Dashboards_ReportCatalog"

# Where the published catalog lives. Same two-host arrangement as
# TPxi_OpsCheckList: the public domain is fine from a browser, but Cloudflare
# challenges TouchPoint's own egress, so server-side fetches use workers.dev.
DC_CATALOG_PUBLIC = "https://scripts.displaycache.com/api/report-catalog"
# The marketplace of operational checks. Separate catalog, separate shape: each
# item is a check with its own SQL, a threshold, and the steps to fix what it
# finds. Only the "auto" ones have SQL; "manual" ones are checklist items and
# belong in the checklists tool, not on a dashboard.
DC_OPS_PUBLIC = "https://scripts.displaycache.com/api/ops-catalog"
DC_OPS_WORKER = ("https://touchpoint-scripts.bswaby.workers.dev"
                 "/api/ops-catalog")
DC_CATALOG_WORKER = ("https://touchpoint-scripts.bswaby.workers.dev"
                     "/api/report-catalog")

# Special Content keys. Shared dashboards are visible to everyone the roles
# allow; personal ones belong to one person and are never listed for others.
CONTENT_SHARED = "TPxi_Dashboards_Shared"
CONTENT_USER_PREFIX = "TPxi_Dashboards_User_"
# Before current_user_id was fixed, every user resolved to 0 and so shared one
# storage row. That row is quarantined: never read as anybody's dashboards,
# and only reachable through the admin adoption step below.
LEGACY_UID = 0

GRIDSTACK_CSS = "https://cdn.jsdelivr.net/npm/gridstack@10.3.1/dist/gridstack.min.css"
# The base stylesheet only lays out a 12-column grid. Calling column(6) puts a
# gs-6 class on the container for which no width rules exist, so every tile
# collapses to nothing: the page looks empty at any width below full. This file
# carries the rules for 2 through 11.
GRIDSTACK_EXTRA_CSS = ("https://cdn.jsdelivr.net/npm/gridstack@10.3.1/dist/"
                       "gridstack-extra.min.css")
GRIDSTACK_JS = "https://cdn.jsdelivr.net/npm/gridstack@10.3.1/dist/gridstack-all.js"
CHARTJS = "https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"

# TouchPoint's own home-page widgets are drawn with Google Charts: each one
# calls google.charts.load(...) and google.charts.setOnLoadCallback(...) in its
# own markup. Those calls need google.charts to EXIST, which the loader below
# provides. Without it the widget's script throws on its first line, so the
# widget's shell renders and the chart never appears -- which is exactly how a
# widget tile failed: a title bar and empty space, no error.
GOOGLE_CHARTS = "https://www.gstatic.com/charts/loader.js"

# Roles allowed to create or edit SHARED dashboards. Personal dashboards need
# no special role: they are only ever visible to their owner.
SHARED_EDIT_ROLES = ["Admin", "Developer", "ManageApplication"]
# Managing SOMEONE ELSE'S personal dashboard is a narrower privilege than
# editing a shared one, so it is not simply SHARED_EDIT_ROLES.
ADMIN_MANAGE_ROLES = ["Admin", "Developer"]

# Who may see money. Contribution amounts, donor names and payment balances are
# not ordinary church data: TouchPoint keeps them behind their own roles, and a
# dashboard that reads the same tables has to honour the same line. Enforced
# where the SQL RUNS, not just where tiles are offered -- hiding a tile stops
# nobody who can post an action by hand.
#
# Checked against this install: Finance, FinanceAdmin, FinanceDataEntry and
# FinanceViewOnlyDetail all exist as roles. Enterprise Reporting's own list
# names "ManageFinance", which is NOT a role here, so that entry protects
# nothing; it is left out rather than copied.
CATEGORY_ROLES = {
    "financial": ["Finance", "FinanceAdmin", "FinanceDataEntry",
                  "FinanceViewOnlyDetail", "Admin"],
    # "transactions" is deliberately NOT here. Those reports cover registration
    # fees and outstanding balances on involvements, which is not giving:
    # TouchPoint shows that to the people running the event, and the staff who
    # chase an unpaid camp balance are rarely on the finance team. Contribution
    # data is the line, and that is "financial" above.
    "admin": ["Admin", "Developer"],
}

# Custom tiles are built from a domain rather than a report category, so the
# giving domain maps onto the same rule.
DOMAIN_CATEGORY = {"giving": "financial"}

model.Header = "Dashboards"

# What this deployed copy is, and the key it is published under. The update
# check itself lives in TPxi_Lib_Update so there is one implementation rather
# than a copy per script.
APP_VERSION = "1.1.5"
DC_SCRIPT_ID = "TPxi_Dashboards"
# Two hosts on purpose. The browser checks the version against the public
# domain; the SERVER fetches the code from the workers.dev mirror, because
# Cloudflare challenges requests coming from TouchPoint's own egress.
DC_API_WORKER = "https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint"
# Roles that may overwrite this script's own source.
UPDATE_ROLES = ["Admin", "Developer"]


def can_apply_update():
    for r in UPDATE_ROLES:
        try:
            if model.UserIsInRole(r):
                return True
        except Exception:
            pass
    return False


def get_script_name():
    """What TouchPoint actually installed this as.

    An admin may have named the content slot something other than the
    published id, and writing to the wrong slot would install the new code
    beside the running one instead of over it.
    """
    try:
        posted = get_param("script_name", "").strip()
        if posted:
            return posted
    except Exception:
        pass
    try:
        m = re.search(r"/PyScript(?:Form)?/([^/?#&]+)",
                      str(getattr(model, "URL", "") or ""))
        if m:
            return m.group(1)
    except Exception:
        pass
    return DC_SCRIPT_ID


# ---------------------------------------------------------------------------
# JSON. json.dumps is broken on non-ASCII in IronPython, and every string here
# is text a human typed: dashboard names, tab names, tile titles.
# ---------------------------------------------------------------------------
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
        return '{' + ','.join(_json_escape_string(unicode(k)) + ':' + _json_encode(v)
                              for k, v in obj.items()) + '}'
    if isinstance(obj, (list, tuple)):
        return '[' + ','.join(_json_encode(x) for x in obj) + ']'
    return _json_escape_string(unicode(obj))


def safe_json(obj):
    try:
        return _json_encode(obj)
    except Exception as e:
        return '{"success":false,"error":' + _json_escape_string('safe_json: ' + repr(e)) + '}'


def esc(s):
    if s is None:
        return ""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def get_param(name, default=''):
    try:
        if hasattr(Data, name):
            v = getattr(Data, name)
            if v is not None:
                return str(v)
    except Exception:
        pass
    return default


def decode_payload(s):
    """Undo the entity encoding the browser applies before POSTing.

    ASP.NET request validation rejects the entire POST if any value looks like
    markup, and a tile title of "<3 our guests" is enough to trip it. The client
    encodes the angle brackets; we put them back before parsing.
    """
    if not s:
        return s
    return (s.replace('&lt;', '<').replace('&gt;', '>')
             .replace('&quot;', '"').replace('&#39;', "'")
             .replace('&amp;', '&'))


# ---------------------------------------------------------------------------
# STORAGE
# ---------------------------------------------------------------------------
def load_content_json(name, default):
    try:
        raw = model.TextContent(name)
    except Exception:
        raw = None
    if not raw:
        return default
    try:
        return json.loads(raw)
    except Exception:
        return default


def save_content_json(name, obj):
    model.WriteContentText(name, safe_json(obj), "")


def current_user_id():
    """The signed-in person, or 0 when that cannot be established.

    model.UserPeopleId is a PROPERTY returning int?, not a method. Reading it
    is correct; calling it raises. Callers must treat 0 as "no identity" and
    refuse personal storage, because 0 is also the row every user wrote to
    while this was broken.
    """
    uid = None
    try:
        uid = model.UserPeopleId
    except Exception:
        uid = None
    # Defensive only: works either way if a future build changes the shape.
    if callable(uid):
        try:
            uid = uid()
        except Exception:
            uid = None
    try:
        return int(uid or 0)
    except Exception:
        return 0


def inline_update_js():
    """The version check, without needing TPxi_Lib_Update installed.

    No backslash escapes in the JS below: it lives in a Python string, which
    turns a lone backslash-n into a real newline and ends the literal early.
    """
    return ("""
(function(){
  var SID = "TPXI_SID", MINE = "TPXI_VER", URL = "TPXI_URL";
  function nums(v){
    var p = String(v || "").split("."), o = [];
    for (var i = 0; i < p.length; i++){
      var n = parseInt(p[i], 10);
      o.push(isNaN(n) ? 0 : n);
    }
    return o;
  }
  function newer(a, b){
    var x = nums(a), y = nums(b), len = Math.max(x.length, y.length);
    for (var i = 0; i < len; i++){
      var xa = x[i] || 0, yb = y[i] || 0;
      if (xa !== yb) return xa > yb;
    }
    return false;
  }
  function show(latest){
    var el = document.getElementById("tpxiUpdate");
    if (!el){
      el = document.createElement("div");
      el.id = "tpxiUpdate";
      if (document.body) document.body.insertBefore(el, document.body.firstChild);
    }
    var canApply = (typeof window.applyAppUpdate === "function");
    el.innerHTML =
      "<div style='margin:8px 0;padding:8px 12px;border:1px solid #b8daff;"
      + "background:#e7f1ff;border-radius:4px;font-size:13px;'>"
      + "A newer version of this script is published: <b>v" + latest
      + "</b>. You are running v" + MINE + ". "
      + (canApply
         ? "<button type='button' id='tpxiUpdateBtn' style='margin-left:6px;"
           + "padding:4px 10px;background:#0078d4;color:#fff;border:0;"
           + "border-radius:3px;cursor:pointer;'>Update now</button>"
         : "<a href='https://scripts.displaycache.com' target='_blank'"
           + " rel='noopener'>Get it</a>")
      + "</div>";
    if (canApply){
      var b = document.getElementById("tpxiUpdateBtn");
      if (b) b.onclick = function(){ window.applyAppUpdate(); };
    }
  }
  try {
    var xhr = new XMLHttpRequest();
    // On the query string rather than a header: a custom header would turn
    // this into a preflighted request for no gain. The value is the site the
    // page is already served from, which the browser sends as Origin anyway.
    var u = URL + "?s=" + encodeURIComponent(SID) + "&v=" + encodeURIComponent(MINE);
    xhr.open("GET", u, true);
    xhr.timeout = 5000;
    xhr.onreadystatechange = function(){
      if (xhr.readyState !== 4 || xhr.status !== 200) return;
      try {
        var v = JSON.parse(xhr.responseText)[SID];
        if (v && newer(v, MINE)) show(v);
      } catch (e) { }
    };
    xhr.send();
  } catch (e) { }
})();
"""
            .replace("TPXI_SID", DC_SCRIPT_ID)
            .replace("TPXI_VER", APP_VERSION)
            .replace("TPXI_URL", "https://scripts.displaycache.com"
                                 "/api/touchpoint/script-versions"))


def dc_headers(extra=None):
    """Headers for a call out to the catalog.

    Server-side calls carry no identity of their own, so without this the
    catalog cannot tell one church from another or say who is running an old
    version. Same convention TPxi Go already uses: the host, the script, the
    version. No person, no PeopleId, nothing about the database.

    Set the admin setting TPxiSendHost to 0 to send nothing but Accept.
    """
    h = {"Accept": "application/json"}
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
    h["X-TPxi-Version"] = APP_VERSION
    if extra:
        h.update(extra)
    return h


def my_roles():
    """Role names this person holds."""
    out = []
    try:
        for r in q.QuerySql(
                "SELECT ro.RoleName FROM dbo.Roles ro WITH (NOLOCK) "
                "JOIN dbo.UserRole ur WITH (NOLOCK) ON ur.RoleId = ro.RoleId "
                "JOIN dbo.Users u WITH (NOLOCK) ON u.UserId = ur.UserId "
                "WHERE u.PeopleId = %d" % int(current_user_id() or 0)):
            nm = str(getattr(r, "RoleName", "") or "")
            if nm and nm not in out:
                out.append(nm)
    except Exception:
        pass
    return out


def can_manage_users():
    """May act on another person's personal dashboards."""
    for r in ADMIN_MANAGE_ROLES:
        try:
            if model.UserIsInRole(r):
                return True
        except Exception:
            pass
    return False


def acting_uid(uid):
    """Whose dashboards this request is really about.

    An admin can pass owner_uid to work inside someone else's storage. Checked
    on every request rather than trusted from the page: the client having shown
    an admin screen is not proof the caller is an admin.
    """
    other = get_param("owner_uid", "").strip()
    if other and other.isdigit() and int(other) != uid and can_manage_users():
        return int(other)
    return uid


def all_dashboard_owners():
    """Everyone who has personal dashboards stored, newest slots first.

    Each person's dashboards live in one content row named for their PeopleId,
    so the list comes from the row names rather than from a registry that
    could drift out of step with what is actually stored.
    """
    rows = []
    try:
        rows = q.QuerySql("""
            SELECT c.Name,
                   ISNULL(p.Name2, '') AS OwnerName,
                   ISNULL(p.EmailAddress, '') AS OwnerEmail
            FROM dbo.Content c WITH (NOLOCK)
            LEFT JOIN dbo.People p WITH (NOLOCK)
                   ON TRY_CAST(REPLACE(c.Name, 'TPxi_Dashboards_User_', '')
                               AS INT) = p.PeopleId
            WHERE c.Name LIKE 'TPxi_Dashboards_User_%'
            ORDER BY ISNULL(p.Name2, c.Name)
        """) or []
    except Exception:
        rows = []
    out = []
    for r in rows:
        name = str(getattr(r, "Name", "") or "")
        tail = name.replace(CONTENT_USER_PREFIX, "")
        if not tail.isdigit() or int(tail) == LEGACY_UID:
            continue
        out.append({"uid": int(tail),
                    "owner": str(getattr(r, "OwnerName", "") or "") or ("PeopleId " + tail),
                    "email": str(getattr(r, "OwnerEmail", "") or "")})
    return out


def people_names(ids):
    """PeopleId -> name, for ids that did not come from a dashboard slot."""
    out = {}
    clean = [str(int(i)) for i in ids if str(i).strip().isdigit()]
    if not clean:
        return out
    try:
        rows = q.QuerySql(
            "SELECT PeopleId, ISNULL(Name2, '') AS Nm FROM dbo.People WITH (NOLOCK) "
            "WHERE PeopleId IN (%s)" % ",".join(clean[:500])) or []
        for r in rows:
            out[int(r.PeopleId)] = str(getattr(r, "Nm", "") or "")
    except Exception:
        pass
    return out


def can_edit_shared():
    for r in SHARED_EDIT_ROLES:
        try:
            if model.UserIsInRole(r):
                return True
        except Exception:
            pass
    return False


# Settings a report can depend on: label, where the value comes from, and the
# fallback. A report names these as {tokens} in its SQL, so the catalog already
# knows which ones each report uses; this turns that into something a person
# can read and check before trusting a number.
SETTING_TOKENS = {
    "serving_types": ("What counts as serving", "serving_member_types",
                      "140,310,320,710"),
    "leader_types": ("What counts as a leader", "leader_member_types",
                     "140,310,320"),
    "active_status_types": ("Which member statuses are active",
                            "active_status_types", "10,20,30"),
    "bg_check_days": ("How long a background check stays valid (days)",
                      "bg_check_days", "730"),
    "fiscal_month": ("Fiscal year start month", "fiscal_year_start_month", "10"),
    "waag_filter": ("Which involvements count as attendance",
                    "attendance_scope", "waag"),
    "bg_approval_ok": ("Which background check approvals you accept",
                       "bg_approval_ok", "Approved"),
}

# Values BackgroundChecks.ApprovalStatus actually holds, from
# lookup.BackGroundCheckApprovalCodes. Stored as the description because that
# is what the column contains, not the code.
BG_APPROVALS = ["Approved", "Pending", "Not Approved", "Abandoned"]


def bg_approval_list():
    """(values, was_set). Falls back to Approved, and says that it guessed.

    Left unconfigured this decides whether someone counts as cleared to be
    with minors, so the report says out loud that nobody has confirmed it
    rather than looking authoritative on a default.
    """
    raw = str((load_settings() or {}).get("bg_approval_ok", "") or "").strip()
    if not raw:
        return ["Approved"], False
    vals = [x.strip() for x in raw.split(",") if x.strip() in BG_APPROVALS]
    if not vals:
        return ["Approved"], False
    return vals, True

# TouchPoint's own Week At A Glance decides what counts by whether a division
# has a ReportLine. That is an editorial choice the church already made, and
# leaving it out is why a dashboard can disagree with the report staff trust:
# here it is 8,624 a week against 9,380 for every meeting in the database. The
# difference is real ministry -- choir rehearsals, Peer Place, enrichment --
# deliberately kept off the weekly report, so this stays switchable.
# Two gates, per TouchPoint's spec: a Program appears when it has a RptGroup,
# and a Division appears as a row under it when it has a RptLine (the number
# is the display order, so presence is the inclusion test). Both are checked
# even though every RptLine division here currently sits under a RptGroup
# program -- a future ReportLine added under an unlisted program would
# otherwise count here and not on the report it is meant to match.
# https://docs.touchpointsoftware.com/SummaryReports/WeekAtAGlanceSpecs.html
WAAG_SQL = ("AND EXISTS (SELECT 1 FROM dbo.DivOrg wdo WITH (NOLOCK) "
            "JOIN dbo.Division wdv WITH (NOLOCK) ON wdv.Id = wdo.DivId "
            "JOIN dbo.Program wpg WITH (NOLOCK) ON wpg.Id = wdv.ProgId "
            "WHERE wdo.OrgId = o.OrganizationId "
            "AND ISNULL(wdv.ReportLine, '') <> '' "
            "AND wpg.RptGroup IS NOT NULL)")


def sql_tokens(sql):
    """Tokens a piece of SQL actually contains."""
    return sorted(set(re.findall(r"\{([a-z_0-9]+)\}", sql or "")))


def setting_notes(tokens):
    """(label, current value) for each setting a report depends on."""
    st = load_settings() or {}
    out = []
    for t in (tokens or []):
        info = SETTING_TOKENS.get(t)
        if not info:
            continue
        label, key, fallback = info
        val = st.get(key)
        out.append({"token": t, "label": label,
                    "value": str(val if val not in (None, "") else fallback),
                    "is_default": val in (None, "")})
    return out


def category_roles(cat):
    """Roles that may see a category. Overridable per church in settings."""
    st = load_settings() or {}
    override = st.get("category_roles") or {}
    if cat in override:
        return [r for r in (override[cat] or []) if str(r).strip()]
    return CATEGORY_ROLES.get(cat, [])


def can_view_category(cat):
    roles = category_roles(cat)
    if not roles:
        return True                      # unrestricted category
    for r in roles:
        try:
            if model.UserIsInRole(r):
                return True
        except Exception:
            pass
    return False


def load_settings():
    d = load_content_json(CONTENT_SETTINGS, {})
    return d if isinstance(d, dict) else {}


def save_settings(d):
    save_content_json(CONTENT_SETTINGS, d)


_REPORTS_SCRIPT_CACHE = [None]


def reports_script_name():
    """Which script serves report tiles, if any.

    Order: an explicit setting, then the compiled-in default, then a search of
    the installed Python scripts for one that answers run_report. Returns "" if
    nothing plausible is installed, and callers must cope with that rather than
    assuming reports exist.
    """
    if _REPORTS_SCRIPT_CACHE[0] is not None:
        return _REPORTS_SCRIPT_CACHE[0]
    name = str((load_settings() or {}).get("reports_script", "") or "").strip()
    if not name:
        name = REPORTS_SCRIPT_DEFAULT
    if not name:
        try:
            rows = q.QuerySql("""
                SELECT TOP 3 Name FROM dbo.Content WITH (NOLOCK)
                WHERE TypeID = 5 AND Archived IS NULL
                  AND CAST(Body AS VARCHAR(MAX)) LIKE '%def get_default_reports%'
                  AND CAST(Body AS VARCHAR(MAX)) LIKE '%run_report%'
                ORDER BY LEN(Name)
            """)
            for r in rows:
                nm = str(getattr(r, "Name", "") or "").strip()
                if nm:
                    name = nm
                    break
        except Exception:
            name = ""
    _REPORTS_SCRIPT_CACHE[0] = name or ""
    return _REPORTS_SCRIPT_CACHE[0]


def load_shared():
    d = load_content_json(CONTENT_SHARED, {"dashboards": []})
    if not isinstance(d, dict):
        d = {"dashboards": []}
    d.setdefault("dashboards", [])
    return d


def load_personal(uid):
    # uid 0 is both "we could not tell who you are" and the legacy row that
    # every user shared before current_user_id was fixed. Serving it would
    # hand one person's dashboards to the next, so it is never served here.
    # legacy_personal() is the one deliberate reader, for admin migration.
    if not uid:
        return {"dashboards": []}
    d = load_content_json(CONTENT_USER_PREFIX + str(uid), {"dashboards": []})
    if not isinstance(d, dict):
        d = {"dashboards": []}
    d.setdefault("dashboards", [])
    return d


def legacy_personal():
    """The quarantined row, read only so an admin can adopt or discard it."""
    d = load_content_json(CONTENT_USER_PREFIX + str(LEGACY_UID),
                          {"dashboards": []})
    if not isinstance(d, dict):
        d = {"dashboards": []}
    d.setdefault("dashboards", [])
    return d


def legacy_hints(dash):
    """Who probably built this one.

    Nothing was recorded, so this infers. A tile scoped to a Search Builder
    search names that search, and dbo.Query knows who owns it. That is a
    strong hint and nothing more, so it is shown as evidence for a person to
    judge rather than used to assign anything automatically.
    """
    names = set()
    for tab in (dash.get("tabs") or []):
        for t in (tab.get("tiles") or []):
            qn = str(t.get("query", "") or "").strip()
            if qn:
                names.add(qn)
    if not names:
        return []
    safe = [x.replace("'", "''") for x in list(names)[:25]]
    try:
        rows = q.QuerySql("""
            SELECT DISTINCT qq.name AS SearchName,
                   ISNULL(p.Name2, qq.owner) AS OwnerName,
                   ISNULL(u.PeopleId, 0) AS OwnerUid
            FROM dbo.Query qq WITH (NOLOCK)
            LEFT JOIN dbo.Users u WITH (NOLOCK) ON u.Username = qq.owner
            LEFT JOIN dbo.People p WITH (NOLOCK) ON p.PeopleId = u.PeopleId
            WHERE qq.name IN ('%s')
        """ % "','".join(safe)) or []
    except Exception:
        return []
    out = []
    for r in rows:
        out.append({"search": str(getattr(r, "SearchName", "") or ""),
                    "owner": str(getattr(r, "OwnerName", "") or ""),
                    "uid": int(getattr(r, "OwnerUid", 0) or 0)})
    return out


def legacy_rows():
    """The quarantined dashboards, one row each, with what is known about them."""
    out = []
    for i, d in enumerate(legacy_personal().get("dashboards", [])):
        tiles = 0
        for tab in (d.get("tabs") or []):
            tiles += len(tab.get("tiles") or [])
        out.append({"idx": i,
                    "id": str(d.get("id", "") or ""),
                    "name": str(d.get("name", "") or "(untitled)"),
                    "tabs": len(d.get("tabs") or []),
                    "tiles": tiles,
                    "hints": legacy_hints(d)})
    return out


def adopt_many(idxs, to_uid):
    """Give one or more quarantined dashboards to one person.

    Matched on position rather than id because ids were handed out inside the
    old shared row and are not unique across it. Positions shift as entries
    leave, so the whole batch is taken in one pass instead of one at a time.
    """
    if not can_manage_users():
        return False, "You do not have permission to do that."
    if not to_uid:
        return False, "Choose who these belong to."
    store = legacy_personal()
    items = store.get("dashboards", [])
    want = sorted(set(i for i in idxs if 0 <= i < len(items)))
    if not want:
        return False, "Those are no longer there. Reload the list."
    mine = load_personal(to_uid)
    lst = mine.get("dashboards", [])
    taken = set(str(x.get("id")) for x in lst)
    for i in want:
        d = dict(items[i])
        d["owner"] = to_uid
        d["scope"] = "personal"
        if str(d.get("id")) in taken:
            d["id"] = next_id(lst, scope_prefix("personal"))
        taken.add(str(d.get("id")))
        lst.append(d)
    mine["dashboards"] = lst
    save_content_json(CONTENT_USER_PREFIX + str(to_uid), mine)
    # Removed only after the destination write has gone through, so a failure
    # leaves the quarantined copies in place rather than losing them.
    store["dashboards"] = [d for i, d in enumerate(items) if i not in set(want)]
    save_content_json(CONTENT_USER_PREFIX + str(LEGACY_UID), store)
    who = people_names([to_uid]).get(to_uid, "them")
    return True, "Moved %d to %s." % (len(want), who)


def reassign_dashboards(refs, to_uid):
    """Hand dashboards that already have an owner to somebody else.

    refs are "uid:scope:id". A personal dashboard physically moves between
    storage rows; a shared one stays shared and only its owner stamp changes,
    since sharing is what makes it visible, not ownership.
    """
    if not can_manage_users():
        return False, "You do not have permission to do that."
    if not to_uid:
        return False, "Choose who these should belong to."
    dest = load_personal(to_uid)
    dest_lst = dest.get("dashboards", [])
    taken = set(str(x.get("id")) for x in dest_lst)
    # Source uid -> its list as it stands MID-BATCH. Re-reading storage per ref
    # would undo an earlier removal from the same owner, moving a dashboard
    # while also leaving it behind.
    working = {}
    moved = 0
    shared_touched = False
    shared = None
    for ref in refs:
        bits = str(ref).split(":")
        if len(bits) != 3:
            continue
        src_s, scope, did = bits[0], bits[1], bits[2]
        if scope == "shared":
            if shared is None:
                shared = load_shared()
            for d in shared.get("dashboards", []):
                if str(d.get("id")) == did:
                    d["owner"] = to_uid
                    moved += 1
                    shared_touched = True
            continue
        if not src_s.isdigit():
            continue
        src = int(src_s)
        if src == to_uid:
            continue                      # already theirs
        if src not in working:
            working[src] = load_personal(src).get("dashboards", [])
        keep, took = [], None
        for d in working[src]:
            if took is None and str(d.get("id")) == did:
                took = d
            else:
                keep.append(d)
        if took is None:
            continue
        took = dict(took)
        took["owner"] = to_uid
        took["scope"] = "personal"
        if str(took.get("id")) in taken:
            took["id"] = next_id(dest_lst, scope_prefix("personal"))
        taken.add(str(took.get("id")))
        dest_lst.append(took)
        working[src] = keep
        moved += 1
    if not moved:
        return False, "Nothing moved. Reload the list and try again."
    if dest_lst:
        dest["dashboards"] = dest_lst
        save_content_json(CONTENT_USER_PREFIX + str(to_uid), dest)
    # Sources are emptied only after the destination is written, so an
    # interruption duplicates rather than loses.
    for src in working:
        st = load_personal(src)
        st["dashboards"] = working[src]
        save_content_json(CONTENT_USER_PREFIX + str(src), st)
    if shared_touched and shared is not None:
        save_content_json(CONTENT_SHARED, shared)
    who = people_names([to_uid]).get(to_uid, "them")
    return True, "Moved %d dashboard(s) to %s." % (moved, who)


def drop_one(idx):
    if not can_manage_users():
        return False, "You do not have permission to do that."
    store = legacy_personal()
    items = store.get("dashboards", [])
    if idx < 0 or idx >= len(items):
        return False, "That one is no longer there. Reload the list."
    nm = str(items[idx].get("name", "") or "(untitled)")
    del items[idx]
    store["dashboards"] = items
    save_content_json(CONTENT_USER_PREFIX + str(LEGACY_UID), store)
    return True, "Deleted %s." % nm


def user_can_see(dash):
    """Roles on a dashboard. Empty list means everyone, matching how TouchPoint
    treats DashboardWidgetRoles and how Enterprise Reporting treats categories.
    """
    roles = dash.get("roles") or []
    if not roles:
        return True
    for r in roles:
        try:
            if model.UserIsInRole(str(r)):
                return True
        except Exception:
            pass
    return False


def visible_dashboards(uid):
    """Everything this person may open: their own, plus shared ones they pass
    the role check for. Personal dashboards are never exposed to anyone else.
    """
    out = []
    for d in load_personal(uid).get("dashboards", []):
        d = dict(d)
        d["scope"] = "personal"
        out.append(d)
    for d in load_shared().get("dashboards", []):
        if user_can_see(d):
            d = dict(d)
            d["scope"] = "shared"
            out.append(d)
    out.sort(key=lambda x: (x.get("name") or "").lower())
    return out


def find_dashboard(uid, did, scope=None):
    vis = visible_dashboards(uid)
    if scope:
        for d in vis:
            if str(d.get("id")) == str(did) and d.get("scope") == scope:
                return d
    for d in vis:
        if str(d.get("id")) == str(did):
            return d
    return None


def next_id(existing, prefix):
    n = 1
    used = set(str(x.get("id", "")) for x in existing)
    while (prefix + str(n)) in used:
        n += 1
    return prefix + str(n)


def scope_prefix(scope):
    """Ids carry their store, so the two can never collide in a link."""
    return "s" if scope == "shared" else "p"


def store_dashboard(uid, dash):
    """Write a dashboard into whichever store it belongs to.

    Editing a shared dashboard is gated here rather than in the UI alone, so a
    crafted POST cannot rewrite a shared board from someone without the role.
    """
    scope = dash.get("scope") or "personal"
    if scope == "shared":
        if not can_edit_shared():
            return False, "You do not have permission to edit shared dashboards."
        store = load_shared()
        key = CONTENT_SHARED
    else:
        if not uid:
            return False, ("Could not tell who you are signed in as, so there "
                           "is nowhere private to save this. Sign in again.")
        store = load_personal(uid)
        key = CONTENT_USER_PREFIX + str(uid)
        dash["owner"] = uid

    prev = str(dash.get("prev_scope") or "")
    if prev and prev != scope and dash.get("id"):
        if prev == "shared" and not can_edit_shared():
            return False, "You do not have permission to move a shared dashboard."
        old_store = load_shared() if prev == "shared" else load_personal(uid)
        old_key = (CONTENT_SHARED if prev == "shared"
                   else CONTENT_USER_PREFIX + str(uid))
        old_store["dashboards"] = [d for d in old_store.get("dashboards", [])
                                   if str(d.get("id")) != str(dash.get("id"))]
        save_content_json(old_key, old_store)
        # A new home means a new id namespace; keeping the old one risks a
        # collision with an existing entry in the destination store.
        dash["id"] = ""
    dash.pop("prev_scope", None)

    lst = store.get("dashboards", [])
    if not dash.get("id"):
        dash["id"] = next_id(lst, scope_prefix(scope))
    replaced = False
    for i, d in enumerate(lst):
        if str(d.get("id")) == str(dash.get("id")):
            lst[i] = dash
            replaced = True
            break
    if not replaced:
        lst.append(dash)
    store["dashboards"] = lst
    save_content_json(key, store)
    return True, dash["id"]


def delete_dashboard(uid, did, scope):
    if scope == "shared":
        if not can_edit_shared():
            return False, "You do not have permission to delete shared dashboards."
        store, key = load_shared(), CONTENT_SHARED
    else:
        if not uid:
            return False, "Could not tell who you are signed in as."
        store, key = load_personal(uid), CONTENT_USER_PREFIX + str(uid)
    before = len(store.get("dashboards", []))
    store["dashboards"] = [d for d in store.get("dashboards", [])
                           if str(d.get("id")) != str(did)]
    if len(store["dashboards"]) == before:
        return False, "Not found."
    save_content_json(key, store)
    return True, "Deleted."


# ---------------------------------------------------------------------------
# LIBRARY
# Curated starting points, referencing real Enterprise Reporting report ids.
# "Add" clones one into the person's own space so it is theirs to change; the
# template is never edited in place.
# ---------------------------------------------------------------------------
# A table tile shows a fixed window and scrolls inside it. It used to be
# marked to grow to its content, which the Add-tile path never did: a report
# returning a couple of hundred rows stretched the tile to the height of the
# whole list and pushed everything below it off the screen.
def _t(rid, title, disp, x, y, w, h, agg=None):
    """A report tile in a template.

    agg lets a template show a plain list as a number or a chart -- {"fn":
    "count"} for "how many are in this list", or a group-by with a measure.
    Without it a chart tile falls back to whatever the report declares, which
    is nothing for the many reports that are only ever tables.
    """
    if disp == "table":
        # Six units, because header, action bar and column headings eat most of
        # a 4-unit tile and leave room for about three rows. Kept inline rather
        # than as a module constant: the catalog exporter slices this file from
        # this def, so anything above it is invisible to that build.
        h = max(h, 6)
    t = {"kind": "report", "report_id": rid, "title": title,
         "display": disp, "x": x, "y": y, "w": w, "h": h, "filters": {},
         "fit": (disp not in ("chart", "table"))}
    if agg:
        base = {"label": "", "fn": "count", "value": "", "chart": "bar",
                "top": "15"}
        base.update(agg)
        t["agg"] = base
    return t


def _stack(tabs):
    """Re-stack each tab's rows so raised table heights cannot overlap.

    Rows are taken from the authored y, so the arrangement the template
    intended is preserved exactly; only the vertical offsets are recomputed.
    """
    for tab in (tabs or []):
        rows = {}
        for t in (tab.get("tiles") or []):
            rows.setdefault(t.get("y", 0), []).append(t)
        y = 0
        for key in sorted(rows.keys()):
            row = rows[key]
            tall = max(t.get("h", 4) for t in row)
            for t in row:
                t["y"] = y
            y += tall
    return tabs


def _k(check, title, disp, x, y, w, h):
    """A marketplace check tile in a template."""
    if disp == "table":
        h = max(h, 6)
    return {"kind": "opscheck", "check_id": check, "title": title,
            "display": disp, "x": x, "y": y, "w": w, "h": h,
            "fit": (disp == "kpi")}


def _w(name, title, x, y, w, h):
    """A widget tile in a template, referenced by name and resolved on Add."""
    return {"kind": "widget", "widget_name": name, "title": title,
            "x": x, "y": y, "w": w, "h": h, "fit": True}


def resolve_widget_names(tabs):
    """Turn widget_name into the widget_id on THIS database.

    Returns (tabs, missing_names). A template that names a widget this church
    does not have drops that tile rather than shipping a dead one, and says
    which so the result is explainable.
    """
    by_name = {}
    try:
        # Names are not unique: this database has both "Giving Sources" and
        # "Giving Sources " (trailing space), which normalise to the same key
        # but carry DIFFERENT roles. Without an order the winner was whichever
        # row the engine happened to return first, so a template could bind to
        # the copy the viewer has no role for and render nothing. Ordered so
        # the pick is deterministic and prefers a widget that has roles at all,
        # since one with none can never be embedded.
        for r in q.QuerySql("""
                SELECT w.Id, w.Name
                FROM dbo.DashboardWidgets w WITH (NOLOCK)
                ORDER BY CASE WHEN EXISTS (
                             SELECT 1 FROM dbo.DashboardWidgetRoles dr
                             WHERE dr.WidgetId = w.Id) THEN 0 ELSE 1 END,
                         w.Enabled DESC, w.Id
        """):
            nm = str(getattr(r, "Name", "") or "").strip().lower()
            if nm and nm not in by_name:
                by_name[nm] = int(r.Id)
    except Exception:
        return (tabs, [])
    missing = []
    for tab in tabs:
        keep = []
        for t in (tab.get("tiles") or []):
            if t.get("kind") == "widget" and not t.get("widget_id"):
                key = str(t.get("widget_name", "")).strip().lower()
                wid = by_name.get(key)
                if not wid:
                    missing.append(t.get("widget_name", "?"))
                    continue
                t = dict(t)
                t["widget_id"] = wid
                t.pop("widget_name", None)
            keep.append(t)
        tab["tiles"] = keep
    return (tabs, missing)


def library_needs(lib):
    """What a template depends on: 'reports', 'widgets', or neither."""
    need = set()
    for tab in (lib.get("tabs") or []):
        for t in (tab.get("tiles") or []):
            k = t.get("kind")
            if k == "report":
                need.add("reports")
            elif k == "widget":
                need.add("widgets")
    return sorted(need)


def _shape_hash(tabs):
    """A stable fingerprint of a board's LAYOUT and CONTENT.

    Deliberately ignores volatile things (widget ids resolved per database,
    tile source chosen at install) so the same template installed twice on
    different databases fingerprints the same. Two uses: telling whether a
    template has moved on since someone installed it, and telling whether that
    person has since changed their own copy.
    """
    parts = []
    for tab in (tabs or []):
        row = [str(tab.get("name", ""))]
        for t in (tab.get("tiles") or []):
            row.append("|".join([
                str(t.get("kind", "")),
                str(t.get("report_id", "") or t.get("widget_name", "")),
                str(t.get("title", "")),
                str(t.get("display", "")),
                str(t.get("x", "")), str(t.get("y", "")),
                str(t.get("w", "")), str(t.get("h", "")),
                safe_json(t.get("agg") or {}),
                safe_json(t.get("filters") or {}),
            ]))
        parts.append("~".join(row))
    raw = "^".join(parts)
    # No hashlib in this runtime; a rolling checksum is plenty to notice that
    # something changed, which is all this is asked to do.
    h = 0
    for ch in raw:
        h = (h * 131 + ord(ch)) & 0xFFFFFFFF
    return "%08x" % h


def library_version(lib):
    return _shape_hash(lib.get("tabs") or [])


def library_report_ids(lib):
    return sorted(set(t.get("report_id") for tab in (lib.get("tabs") or [])
                      for t in (tab.get("tiles") or [])
                      if t.get("kind") == "report" and t.get("report_id")))


def library_available(lib):
    """(usable, reason). A template is offered only if this install can run it.

    Report tiles are served by the catalog when the definition is available
    there, and only fall back to Enterprise Reporting when it is not. These
    templates predate the catalog and used to demand Enterprise Reporting even
    though every report they name now ships in it.
    """
    need = library_needs(lib)
    if "reports" not in need:
        return (True, "")
    if reports_script_name():
        return (True, "")
    wanted = set(library_report_ids(lib))
    have = set(str(r.get("id", "")) for r in
               load_report_catalog().get("reports", []))
    if wanted and wanted.issubset(have):
        return (True, "")
    published, err = fetch_published_catalog()
    if not err:
        reports = (published.get("reports", []) if isinstance(published, dict)
                   else (published or []))
        if wanted.issubset(have | set(str(r.get("id", "")) for r in reports)):
            return (True, "")
    return (False, "Needs Enterprise Reporting or the report catalog, and "
                   "neither can supply every report this template uses.")


def get_library():
    return [
        {"id": "lib_overview", "name": "Church Overview", "icon": "fa-church",
         "audience": "Anyone, as a home dashboard",
         "description": "The one-screen view: who is here, how often they come, "
                        "and what they give. A good default home dashboard.",
         # Numbers first, then the trends behind them, then the make-up of the
         # congregation. The attendance figures include head-counted services,
         # which is most of Sunday morning.
         "tabs": [
             {"name": "This Week", "tiles": [
                 _t("eng_yoy_snapshot", "Year over Year", "table", 0, 0, 6, 6),
                 _t("att_weekly_trends", "Weekly Attendance", "chart", 6, 0, 6, 6),
                 _t("fin_weekly_giving", "Weekly Giving", "chart", 0, 6, 6, 5),
                 _t("att_by_program", "Attendance by Program", "chart", 6, 6, 6, 5),
             ]},
             {"name": "People", "tiles": [
                 _t("demo_member_status", "Member Status", "chart", 0, 0, 6, 5),
                 _t("demo_age_distribution", "Age Distribution", "chart", 6, 0, 6, 5),
                 _t("eng_score_distribution", "Engagement", "chart", 0, 5, 6, 5),
                 _t("mem_growth_trends", "Membership Growth", "chart", 6, 5, 6, 5),
             ]},
             {"name": "Coming and Going", "tiles": [
                 _t("att_lapsed", "Stopped Coming", "kpi", 0, 0, 4, 3,
                    agg={"fn": "count"}),
                 _t("eng_first_time_guests", "First-Time Guests", "kpi", 4, 0, 4, 3,
                    agg={"fn": "count"}),
                 _t("mem_new_members", "New Members", "kpi", 8, 0, 4, 3,
                    agg={"fn": "count"}),
                 _t("att_lapsed", "Who Stopped Coming", "table", 0, 3, 6, 6),
                 _t("eng_first_time_guests", "Recent First-Time Guests",
                    "table", 6, 3, 6, 6),
             ]}]},

        {"id": "lib_attendance", "name": "Attendance Health", "icon": "fa-users",
         "audience": "Ministry leads",
         "description": "How attendance is moving, which ministries and times "
                        "carry it, and the names to follow up. Built for a "
                        "ministry review meeting.",
         # Attendance figures here include head-counted services, which is most
         # of Sunday morning: counting only individual check-ins reported 1,158
         # against an actual 43,098 for Worship.
         "tabs": [
             {"name": "Trends", "tiles": [
                 _t("att_weekly_trends", "Average Week", "kpi", 0, 0, 3, 3,
                    agg={"fn": "avg", "value": "Attendance",
                         "label": "Attendance in a typical week"}),
                 _t("att_weekly_trends", "People in a Week", "kpi", 3, 0, 3, 3,
                    agg={"fn": "avg", "value": "UniquePeople",
                         "label": "Different people, typical week"}),
                 _t("att_lapsed", "Lapsed Attenders", "kpi", 6, 0, 3, 3,
                    agg={"fn": "count"}),
                 _t("eng_lapsed_children", "Children Not Back", "kpi",
                    9, 0, 3, 3, agg={"fn": "count"}),
                 _t("att_weekly_trends", "Week by Week", "chart", 0, 3, 12, 5),
                 _t("att_trend_monthly", "Month by Month", "chart", 0, 8, 6, 5),
                 _t("att_seasonal_patterns", "Seasonal Shape", "chart",
                    6, 8, 6, 5),
             ]},
             {"name": "Where People Are", "tiles": [
                 _t("att_by_program", "By Program", "chart", 0, 0, 6, 5),
                 _t("att_by_org_type", "By Involvement Type", "chart",
                    6, 0, 6, 5),
                 _t("att_checkin_by_dow", "By Day of Week", "chart", 0, 5, 6, 5),
                 _t("att_by_meeting_time", "By Time of Day", "chart", 6, 5, 6, 5),
                 _t("att_division_comparison", "By Division", "table",
                    0, 10, 12, 6),
                 _t("att_avg_per_org", "Average by Involvement", "table",
                    0, 16, 12, 6),
             ]},
             {"name": "Who to Follow Up", "tiles": [
                 # Names, not counts. Every tile on this tab is a list someone
                 # can work through.
                 _t("att_lapsed", "Stopped Coming", "table", 0, 0, 12, 7),
                 _t("eng_attendance_gaps", "Slipping in a Group", "table",
                    0, 7, 12, 7),
                 _t("eng_lapsed_adults", "Adults Not Back", "table", 0, 14, 6, 7),
                 _t("eng_lapsed_children", "Children Not Back", "table",
                    6, 14, 6, 7),
             ]},
             {"name": "Families", "tiles": [
                 # A child who comes while the parents do not is a different
                 # conversation from a family that has drifted together.
                 _t("att_family_correlation", "Who Comes Together", "chart",
                    0, 0, 5, 5),
                 _t("eng_parents_not_attending", "Children Here, Parents Not",
                    "table", 5, 0, 7, 5),
                 _t("eng_family_engagement_gap", "Split Families", "table",
                    0, 5, 12, 7),
             ]}]},

        {"id": "lib_giving", "name": "Giving Health", "icon": "fa-donate",
         "audience": "Finance and stewardship",
         "description": "Fund totals, year over year, donor segments and giving "
                        "by age. Pairs with the finance team's monthly review.",
         "tabs": [
             {"name": "Overview", "tiles": [
                 _t("fin_yoy_comparison", "Year over Year", "chart", 0, 0, 8, 4),
                 _t("fin_giving_by_fund", "By Fund", "chart", 8, 0, 4, 4),
                 _t("fin_donor_segments", "Donor Segments", "chart", 0, 4, 6, 4),
                 _t("fin_givers_age_bins", "Givers by Age", "chart", 6, 4, 6, 4),
             ]},
             {"name": "Detail", "tiles": [
                 _t("fin_weekly_giving", "Weekly Giving", "table", 0, 0, 12, 6),
             ]}]},

        {"id": "lib_datahealth", "name": "Data Health", "icon": "fa-broom",
         "audience": "Data stewards",
         "description": "Where the gaps are and whether they are closing, with "
                        "the names behind them on their own tab.",
         # Trend first, names second. A wall of everyone missing an email says
         # nothing about whether the problem is growing; the summary and the
         # intake trend do, and the lists are still one click away.
         "tabs": [
             {"name": "Overview", "tiles": [
                 _t("dh_gaps_summary", "Where the Gaps Are", "chart", 0, 0, 6, 4),
                 _t("adm_data_quality_score", "Data Quality Score", "chart", 6, 0, 6, 4),
                 _t("dh_intake_quality_trend", "Intake Quality by Month",
                    "chart", 0, 4, 12, 5),
                 _t("dh_gaps_summary", "Gap Counts", "table", 0, 9, 6, 6),
                 _t("dh_new_records_incomplete", "New Records Missing Data",
                    "table", 6, 9, 6, 6),
             ]},
             {"name": "Best Practice Checks", "tiles": [
                 # From the DisplayCache marketplace. Counts first, so the tab
                 # reads at a glance and a zero is visibly a good result.
                 _k("Duplicate records (name+DOB)", "Duplicates", "kpi",
                    0, 0, 3, 3),
                 _k("Deceased still in active orgs", "Deceased in Orgs", "kpi",
                    3, 0, 3, 3),
                 _k("21+ still listed as Child", "Adults Marked Child", "kpi",
                    6, 0, 3, 3),
                 _k("Previous Member not archived", "Not Archived", "kpi",
                    9, 0, 3, 3),
                 _k("Just Addeds review", "Just Added", "kpi", 0, 3, 3, 3),
                 _k("Drop Date with recent attendance",
                    "Dropped But Attending", "kpi", 3, 3, 3, 3),
                 _k("Primary/Secondary adult under 18", "Adult Under 18", "kpi",
                    6, 3, 3, 3),
                 _k("Single-person families with Child position",
                    "Alone as Child", "kpi", 9, 3, 3, 3),
                 _k("Duplicate records (name+DOB)", "Duplicates to Merge",
                    "table", 0, 6, 12, 6),
                 _k("Deceased still in active orgs",
                    "Deceased Still Enrolled", "table", 0, 12, 6, 6),
                 _k("21+ still listed as Child",
                    "Adults Still Marked Child", "table", 6, 12, 6, 6),
             ]},
             {"name": "Records and Email", "tiles": [
                 _k("Mrs./Mr. gender mismatch", "Title vs Gender", "table",
                    0, 0, 6, 6),
                 _k("Sensitive role audit", "Who Holds Sensitive Roles",
                    "table", 6, 0, 6, 6),
                 _k("Orgs with no Entry Point", "Involvements Without Entry "
                    "Point", "table", 0, 6, 6, 6),
                 _k("Inactive orgs with current members",
                    "Inactive With Members", "table", 6, 6, 6, 6),
                 _k("Bounced / failed emails (last 7 days)",
                    "Bounced This Week", "table", 0, 12, 6, 6),
                 _k("Spam / reputation failures (last 30 days)",
                    "Spam Reports", "table", 6, 12, 6, 6),
             ]},
             {"name": "Archiving", "tiles": [
                 _k("Archive candidates (3.5yr no activity)",
                    "Archive Candidates", "table", 0, 0, 12, 7),
                 _k("Previous Member not archived",
                    "Previous Members Still Active", "table", 0, 7, 6, 6),
                 _k("Deceased records review", "Deceased to Review", "table",
                    6, 7, 6, 6),
             ]},
             {"name": "Who", "tiles": [
                 _t("dh_members_missing_email", "Missing Email", "table", 0, 0, 6, 6),
                 _t("dh_missing_phone", "Missing Phone", "table", 6, 0, 6, 6),
                 _t("dh_missing_address", "Missing Address", "table", 0, 6, 6, 6),
                 _t("dh_missing_birthdate", "Missing Birth Date", "table", 6, 6, 6, 6),
             ]},
             {"name": "Duplicates", "tiles": [
                 _t("dh_possible_duplicates", "Possible Duplicates", "table", 0, 0, 6, 6),
                 _t("dh_duplicates_by_name", "Duplicates by Name", "table", 6, 0, 6, 6),
                 _t("dh_shared_email", "Email Shared Across Families",
                    "table", 0, 6, 12, 6),
             ]}]},

        {"id": "lib_guests", "name": "Guests and Assimilation", "icon": "fa-door-open",
         "audience": "Connections and assimilation",
         "description": "First-time guests, whether they came back, and new "
                        "members. The assimilation funnel end to end.",
         "tabs": [{"name": "Funnel", "tiles": [
             _t("eng_first_time_guests", "First-Time Guests", "table", 0, 0, 6, 4),
             _t("eng_guest_retention", "Guest Retention", "chart", 6, 0, 6, 4),
             _t("mem_new_members", "New Members", "table", 0, 4, 6, 4),
             _t("eng_parents_not_attending", "Parents Not Attending", "table", 6, 4, 6, 4),
         ]}]},

        {"id": "lib_tasks", "name": "Staff and Tasks", "icon": "fa-tasks",
         "audience": "Staff leadership",
         "description": "Workload, overdue items and aging. For a staff "
                        "meeting where you need to see who is buried.",
         # Counts first, then who is carrying what, then what the work is
         # about. Open means Pending or Accepted: counting Archived as open
         # reported 141,588 open tasks here instead of 3,053.
         "tabs": [
             {"name": "Right Now", "tiles": [
                 _t("task_overdue_by_assignee", "People With Overdue Work", "kpi",
                    0, 0, 4, 3, agg={"fn": "count"}),
                 _t("task_overdue_by_assignee", "Overdue Tasks", "kpi",
                    4, 0, 4, 3, agg={"fn": "sum", "value": "OverdueTasks"}),
                 _t("task_staff_workload", "Staff With Open Work", "kpi",
                    8, 0, 4, 3, agg={"fn": "count"}),
                 _t("task_summary_overview", "Open, Done and Declined", "chart",
                    0, 3, 6, 5),
                 _t("task_aging_analysis", "How Old Is the Backlog", "chart",
                    6, 3, 6, 5),
                 _t("task_overdue_by_assignee", "Who Is Behind", "table",
                    0, 8, 12, 6),
             ]},
             {"name": "Keeping Up", "tiles": [
                 _t("task_completion_velocity", "Created vs Completed", "chart",
                    0, 0, 12, 5),
                 _t("task_completion_by_staff", "Completion by Staff", "chart",
                    0, 5, 6, 5),
                 _t("task_staff_workload", "Workload Distribution", "chart",
                    6, 5, 6, 5),
                 _t("task_completion_by_staff", "Completion Detail", "table",
                    0, 10, 12, 6),
             ]},
             {"name": "What It Is About", "tiles": [
                 _t("task_keyword_trends", "Keyword Trends", "chart", 0, 0, 12, 5),
                 _t("task_people_attention", "People Getting the Most Attention",
                    "table", 0, 5, 12, 6),
             ]}]},

        {"id": "lib_comms", "name": "Communications", "icon": "fa-envelope",
         "audience": "Whoever gets asked why the email did not arrive",
         "description": "Whether email is landing, why it is not, and what is "
                        "still waiting to go out.",
         "tabs": [
             {"name": "Is It Landing", "tiles": [
                 _t("comm_email_stats", "Sent and Delivered", "chart", 0, 0, 12, 5),
                 _t("comm_bounce_rate_trend", "Bounce Rate by Day", "chart",
                    0, 5, 6, 5),
                 _t("comm_failure_breakdown", "Delivered vs Failed", "chart",
                    6, 5, 6, 5),
                 _t("comm_email_opens", "Open Rates by Campaign", "table",
                    0, 10, 12, 6),
             ]},
             {"name": "Why It Failed", "tiles": [
                 # A bounced address is a data problem you fix on the person;
                 # a spam report or reputation failure is about the sending and
                 # editing an address will not touch it.
                 _t("comm_domain_reputation", "Failure by Domain", "table",
                    0, 0, 6, 7),
                 _t("comm_failure_by_person", "Failures by Person", "table",
                    6, 0, 6, 7),
                 _t("comm_email_by_program", "By Program", "chart", 0, 7, 12, 5),
             ]},
             {"name": "Waiting to Send", "tiles": [
                 _t("comm_email_queue", "Pending Queue", "kpi", 0, 0, 4, 3,
                    agg={"fn": "count"}),
                 _t("comm_email_queue", "Still Waiting", "table", 0, 3, 12, 7),
                 _t("comm_email_top_senders", "Who Sends the Most", "table",
                    0, 10, 12, 6),
             ]}]},

        {"id": "lib_exec", "name": "Executive Summary", "icon": "fa-binoculars",
         "audience": "Senior staff, elders, board",
         "description": "The handful of numbers leadership actually asks about: "
                        "are we growing, are people coming, are they giving, and "
                        "are they taking next steps.",
         "tabs": [{"name": "Summary", "tiles": [
             _t("eng_yoy_snapshot", "Year over Year", "table", 0, 0, 6, 4),
             _t("mem_growth_trends", "Membership Growth", "chart", 6, 0, 6, 4),
             _t("att_weekly_trends", "Weekly Attendance", "chart", 0, 4, 6, 4),
             _t("fin_annual_summary", "Annual Giving vs Prior Year", "table", 6, 4, 6, 4),
             _t("eng_next_steps_funnel", "Next Steps Funnel", "chart", 0, 8, 6, 4),
             _t("eng_score_distribution", "Engagement Distribution", "chart", 6, 8, 6, 4),
         ]}]},

        {"id": "lib_risk", "name": "Engagement and Risk", "icon": "fa-heart-broken",
         "audience": "Pastoral care, connections",
         "description": "Who is drifting. Built to be worked, not admired: every "
                        "tile here is a list of names someone should contact.",
         # Numbers, then where people fall out, then the names to work. The
         # counting tiles all sit on UNCAPPED reports: At-Risk and Priority
         # Outreach are TOP 200 / TOP 500, so counting their rows would print
         # exactly 200 and 500 forever and look like a real measurement.
         "tabs": [
             {"name": "Health", "tiles": [
                 _t("eng_unconnected_members", "Unconnected", "kpi", 0, 0, 3, 3,
                    agg={"fn": "count"}),
                 _t("eng_volunteer_burnout", "Burnout Risk", "kpi", 3, 0, 3, 3,
                    agg={"fn": "count"}),
                 _t("eng_first_time_guests", "Guests (30 days)", "kpi", 6, 0, 3, 3,
                    agg={"fn": "count"}),
                 _t("eng_recent_drops", "Recent Drops", "kpi", 9, 0, 3, 3,
                    agg={"fn": "count"}),
                 _t("eng_score_distribution", "Engagement Score", "chart",
                    0, 3, 6, 5),
                 _t("eng_journey_stage_distribution", "Journey Stages", "chart",
                    6, 3, 6, 5),
                 _t("eng_next_steps_funnel", "Next Steps Funnel", "chart",
                    0, 8, 6, 5),
                 _t("eng_guest_retention", "Guest Retention", "chart", 6, 8, 6, 5),
             ]},
             {"name": "Where It Breaks", "tiles": [
                 _t("eng_assimilation_velocity", "How Long Each Stage Takes",
                    "chart", 0, 0, 6, 5),
                 _t("eng_lapse_rate_by_involvement", "Lapse Rate by Involvement",
                    "chart", 6, 0, 6, 5),
                 _t("eng_recent_drops", "Recent Drops", "table", 0, 5, 6, 6),
                 _t("eng_dormant_orgs", "Dormant Involvements", "table", 6, 5, 6, 6),
             ]},
             {"name": "Work the List", "tiles": [
                 # Titled with their caps, because the report decides the limit
                 # and a reader cannot see it from the tile otherwise.
                 _t("eng_at_risk_members", "At-Risk Members (top 200)",
                    "table", 0, 0, 6, 6),
                 _t("eng_priority_outreach", "Priority Outreach (top 500)",
                    "table", 6, 0, 6, 6),
                 _t("eng_attendance_gaps", "Attendance Gaps", "table", 0, 6, 6, 6),
                 _t("eng_unconnected_members", "Unconnected Members",
                    "table", 6, 6, 6, 6),
                 _t("eng_family_engagement_gap", "Family Engagement Gap",
                    "table", 0, 12, 6, 6),
                 _t("eng_volunteer_burnout", "Volunteer Burnout Risk",
                    "table", 6, 12, 6, 6),
             ]}]},

        {"id": "lib_groups", "name": "Groups Health", "icon": "fa-people-group",
         "audience": "Small group and discipleship staff",
         "description": "Which groups are thriving, which are coasting, and "
                        "which have quietly stopped meeting.",
         "tabs": [{"name": "Groups", "tiles": [
             _t("eng_group_growth_decline", "Growth vs Decline", "chart", 0, 0, 8, 4),
             _t("eng_group_lifecycle", "Lifecycle Stages", "table", 8, 0, 4, 4),
             _t("eng_dormant_orgs", "Dormant Involvements", "table", 0, 4, 6, 5),
             _t("eng_group_volatility", "Attendance Volatility", "table", 6, 4, 6, 5),
             _t("eng_cohort_retention", "Cohort Retention", "chart", 0, 9, 12, 4),
         ]}]},

        {"id": "lib_volunteers", "name": "Volunteers and Serving", "icon": "fa-hands-helping",
         "audience": "Volunteer coordinators",
         "description": "Who serves, who is doing too much, and who is not "
                        "cleared to be doing it at all.",
         # Safe church first, because it is the part with a deadline. The two
         # lists there answer different questions and are deliberately both
         # present: one finds adults who are PRESENT with minors whatever the
         # roster says, the other finds everyone the roster calls a volunteer.
         "tabs": [
             # Safe Church is ONLY people actually with minors. The roster
             # list is a different question and lives on its own tab: most of
             # it is adult connect-group leaders who never work with children,
             # so mixing them made the urgent list look like noise.
             {"name": "Safe Church", "tiles": [
                 _t("vol_minors_uncleared", "With Minors, Not Cleared", "kpi",
                    0, 0, 4, 3, agg={"fn": "count"}),
                 _t("vol_minors_uncleared", "Times Present While Uncleared",
                    "kpi", 4, 0, 4, 3,
                    agg={"fn": "sum", "value": "TimesPresent"}),
                 _t("vol_minors_uncleared", "Areas They Cover", "kpi",
                    8, 0, 4, 3, agg={"fn": "sum", "value": "MinorAreas"}),
                 _t("vol_minors_uncleared",
                    "Adults With Minors Without a Current Check",
                    "table", 0, 3, 12, 8),
             ]},
             {"name": "Who Serves", "tiles": [
                 _t("eng_volunteer_participation", "Participation", "chart",
                    0, 0, 6, 5),
                 _t("vol_minors_uncleared", "Areas Covered per Adult", "chart",
                    6, 0, 6, 5,
                    agg={"label": "MinorAreas", "fn": "count", "chart": "bar",
                         "top": "10"}),
                 _t("eng_volunteer_burnout", "Burnout Risk", "table", 0, 5, 12, 6),
             ]},
             {"name": "All Roster Volunteers", "tiles": [
                 _t("eng_bg_check_compliance",
                    "Everyone a Roster Calls a Volunteer, Without a Check",
                    "table", 0, 0, 12, 8),
             ]}]},

        {"id": "lib_membership", "name": "Membership Pipeline", "icon": "fa-user-plus",
         "audience": "Membership and next steps",
         "description": "New members, how they joined, baptisms, and whether "
                        "they stay once they are in.",
         "tabs": [{"name": "Pipeline", "tiles": [
             _t("mem_new_members", "New Members", "table", 0, 0, 6, 5),
             _t("mem_join_type", "How They Joined", "chart", 6, 0, 6, 5),
             _t("mem_baptism_stats", "Baptisms", "chart", 0, 5, 6, 4),
             _t("mem_status_transitions", "Status Transitions", "table", 6, 5, 6, 4),
             _t("mem_retention_cohort", "Retention Cohorts", "chart", 0, 9, 12, 4),
         ]}]},

        {"id": "lib_finance_detail", "name": "Giving Detail", "icon": "fa-chart-line",
         "audience": "Finance and stewardship",
         "description": "Beyond the totals: retention, first-time givers, who "
                        "has lapsed, and how people actually give.",
         "tabs": [
             {"name": "Donors", "tiles": [
                 _t("fin_donor_retention", "Donor Retention", "chart", 0, 0, 4, 4),
                 _t("fin_first_time_givers", "First-Time Givers", "table", 4, 0, 8, 4),
                 _t("fin_lapsed_givers", "Lapsed Givers", "table", 0, 4, 12, 5),
             ]},
             {"name": "Patterns", "tiles": [
                 _t("fin_recurring_vs_onetime", "Recurring vs One-Time", "chart", 0, 0, 4, 4),
                 _t("fin_giving_frequency", "Giving Frequency", "chart", 4, 0, 4, 4),
                 _t("fin_giving_by_bundle_type", "By Method", "chart", 8, 0, 4, 4),
                 _t("fin_giving_by_generation", "By Generation", "chart", 0, 4, 6, 4),
                 _t("fin_giving_by_campus", "By Campus", "chart", 6, 4, 6, 4),
             ]}]},

        {"id": "lib_payments", "name": "Payments and Balances", "icon": "fa-file-invoice-dollar",
         "audience": "Registration and finance",
         "description": "Who owes money for events and camps, and how old those "
                        "balances are.",
         "tabs": [{"name": "Balances", "tiles": [
             _t("txn_outstanding_by_program", "Outstanding by Program", "chart", 0, 0, 6, 4),
             _t("txn_aging_report", "Balance Aging", "chart", 6, 0, 6, 4),
             _t("txn_outstanding_by_person", "Outstanding by Person", "table", 0, 4, 12, 5),
             _t("txn_involvements_with_fees", "Involvements with Fees", "table", 0, 9, 12, 4),
         ]}]},

        {"id": "lib_security", "name": "Accounts and Security", "icon": "fa-shield-halved",
         "audience": "Administrators",
         "description": "What needs attention right now, who holds an account, "
                        "and how the system is actually being used.",
         # Three questions in order: is anything wrong, who has access, what
         # are they doing. The numbers across the top are counts of the same
         # lists shown underneath, so a number that looks wrong has its names
         # one tile away. Data quality lives in Data Health, not here.
         "tabs": [
             {"name": "Posture", "tiles": [
                 _t("adm_users_under_attack", "Accounts Under Attack", "kpi",
                    0, 0, 3, 3, agg={"fn": "count"}),
                 _t("adm_locked_accounts", "Locked Out", "kpi",
                    3, 0, 3, 3, agg={"fn": "count"}),
                 # The report is TOP 100, so this counts at most 100. Said in
                 # the title rather than left to read as a total.
                 _t("adm_top_threat_ips", "Threat IPs (top 100)", "kpi",
                    6, 0, 3, 3, agg={"fn": "count"}),
                 _t("adm_stale_accounts", "Stale Accounts", "kpi",
                    9, 0, 3, 3, agg={"fn": "count"}),
                 _t("adm_failed_login_summary", "Failed Logins by Day",
                    "chart", 0, 3, 12, 5),
                 _t("adm_users_under_attack", "Who Is Being Targeted",
                    "table", 0, 8, 6, 6),
                 _t("adm_top_threat_ips", "Where From", "table", 6, 8, 6, 6),
             ]},
             {"name": "Accounts", "tiles": [
                 # Grouped on login recency rather than account status: every
                 # account computes to Active, so that made a one-slice
                 # doughnut. Recency actually spreads.
                 _t("adm_login_activity", "Accounts by Login Recency", "chart",
                    0, 0, 6, 5,
                    agg={"label": "Status", "fn": "count", "chart": "doughnut"}),
                 _t("adm_stale_accounts", "Inactive Longest", "chart", 6, 0, 6, 5,
                    agg={"label": "Name", "fn": "max", "value": "DaysInactive",
                         "chart": "bar", "top": "10"}),
                 _t("adm_stale_accounts", "Stale Accounts", "table", 0, 5, 6, 6),
                 _t("adm_password_resets", "Recent Password Changes",
                    "table", 6, 5, 6, 6),
                 _t("adm_user_accounts", "All User Accounts", "table", 0, 11, 12, 6),
             ]},
             {"name": "Usage", "tiles": [
                 _t("adm_activity_by_type", "What People Do", "chart", 0, 0, 6, 5),
                 _t("adm_mobile_vs_web", "Mobile vs Web", "chart", 6, 0, 6, 5),
                 _t("adm_most_active_users", "Most Active Users", "chart",
                    0, 5, 12, 5),
                 _t("adm_login_activity", "Login Activity", "table", 0, 10, 12, 6),
             ]}]},

        {"id": "lib_techstatus", "name": "Tech Status", "icon": "fa-server",
         "audience": "Whoever keeps TouchPoint running",
         "description": "Logins, security, check-in printing and script "
                        "activity. The morning look before anyone reports a "
                        "problem.",
         "tabs": [
             {"name": "Right Now", "tiles": [
                 _t("adm_users_under_attack", "Accounts Under Attack", "kpi",
                    0, 0, 3, 3, agg={"fn": "count"}),
                 _t("adm_locked_accounts", "Locked Out", "kpi", 3, 0, 3, 3,
                    agg={"fn": "count"}),
                 _t("comm_email_queue", "Email Waiting", "kpi", 6, 0, 3, 3,
                    agg={"fn": "count"}),
                 _t("adm_password_resets", "Recent Resets", "kpi", 9, 0, 3, 3,
                    agg={"fn": "count"}),
                 _t("adm_failed_login_summary", "Failed Logins by Day", "chart",
                    0, 3, 12, 5),
                 _t("adm_users_under_attack", "Who Is Targeted", "table",
                    0, 8, 6, 6),
                 _t("adm_top_threat_ips", "Where From", "table", 6, 8, 6, 6),
             ]},
             {"name": "Printing and Scripts", "tiles": [
                 # Labels are owed to a named kiosk and cleared the moment it
                 # polls, so a backlog is rare and small. One tile covers it;
                 # everything else here is normal activity, because a tab of
                 # exception views is blank on every healthy day.
                 _t("adm_print_queue", "Waiting to Print", "kpi", 0, 0, 4, 3,
                    agg={"fn": "sum", "value": "Waiting",
                         "label": "Labels waiting"}),
                 _t("adm_print_queue", "Backlog by Kiosk", "table", 4, 0, 8, 3),
                 _t("adm_print_recent", "Check-In Printing", "chart", 0, 3, 12, 5),
                 _t("adm_script_runs", "Script Runs by Day", "chart", 0, 8, 6, 5),
                 _t("adm_activity_by_type", "What People Do", "chart", 6, 8, 6, 5),
                 _t("adm_mobile_vs_web", "Mobile vs Web", "chart", 0, 13, 12, 5),
             ]},
             {"name": "Accounts", "tiles": [
                 _t("adm_login_activity", "Accounts by Login Recency", "chart",
                    0, 0, 6, 5,
                    agg={"label": "Status", "fn": "count", "chart": "doughnut"}),
                 _t("adm_most_active_users", "Most Active Users", "chart",
                    6, 0, 6, 5),
                 _t("adm_stale_accounts", "Stale Accounts", "table", 0, 5, 12, 6),
             ]}]},

        {"id": "lib_tp_classics", "name": "TouchPoint Built-Ins", "icon": "fa-puzzle-piece",
         "audience": "Anyone",
         "description": "TouchPoint ships these widgets and most churches never "
                        "switch them on. They are cached by TouchPoint itself, so "
                        "they cost almost nothing to display.",
         "tabs": [{"name": "Built-ins", "tiles": [
             _w("Attendance Year Over Year", "Attendance Year over Year", 0, 0, 6, 5),
             _w("Giving to Budget Comparison", "Giving to Budget", 6, 0, 6, 5),
             _w("Giving Sources", "Giving Sources", 0, 5, 4, 5),
             _w("Division Comparison", "Division Comparison", 4, 5, 4, 5),
             _w("Contact Trend", "Contact Trend", 8, 5, 4, 5),
             _w("Recent Attendance Trends", "Recent Attendance Trends", 0, 10, 6, 5),
             _w("Giving Change Histogram", "Giving Change", 6, 10, 6, 5),
         ]}]},
    ]


def library_status(uid):
    """Per template: what this person has, and whether it is behind.

    Three states matter and are kept apart, because the right action differs:
      up to date   nothing to do
      behind       the template changed and their copy did not: safe to replace
      behind, but they also changed their own copy: replacing loses their work
    A copy installed before versioning has no recorded shape, so it cannot be
    judged and is reported as unknown rather than guessed at.
    """
    out = {}
    mine = visible_dashboards(uid)
    for lib in get_library():
        lid = lib.get("id")
        cur = library_version(lib)
        copies = []
        for d in mine:
            if d.get("from_library") != lid:
                continue
            was = d.get("library_version")
            shape_at_install = d.get("library_installed_shape")
            now_shape = _shape_hash(d.get("tabs") or [])
            if not was or not shape_at_install:
                # Installed before any of this was recorded, so we cannot say
                # WHY it differs. We can still say WHETHER it does: if its
                # shape matches the template exactly there is nothing to do,
                # and if it does not, something moved and it is worth a look.
                # Reporting "unknown" for both was useless, and it hid every
                # template that changed under an older install.
                state = "current" if now_shape == cur else "differs"
            elif was == cur:
                state = "current"
            elif now_shape != shape_at_install:
                state = "behind_edited"
            else:
                state = "behind"
            copies.append({"id": d.get("id"), "name": d.get("name", ""),
                           "scope": d.get("scope", "personal"),
                           "state": state})
        out[lid] = {"version": cur, "copies": copies}
    return out


def clone_from_library(uid, lib_id):
    for lib in get_library():
        if lib.get("id") == lib_id:
            ok, why = library_available(lib)
            if not ok:
                return (False, why)
            tabs, missing = resolve_widget_names(
                json.loads(safe_json(lib.get("tabs", []))))
            _stack(tabs)
            # Point the report tiles at the catalog wherever it can serve them,
            # installing what is missing. Without this a template built one of
            # these boards out of Enterprise Reporting tiles even on an install
            # that has the catalog and does not run that script.
            wanted = library_report_ids(lib)
            served = ensure_catalog_reports(wanted) if wanted else set()
            for tab in tabs:
                for t in (tab.get("tiles") or []):
                    if t.get("kind") == "report" and t.get("report_id") in served:
                        t["source"] = "catalog"
            dash = {
                "id": "",
                "name": lib.get("name", "Dashboard"),
                "description": lib.get("description", ""),
                "icon": lib.get("icon", "fa-th-large"),
                "roles": [],
                "scope": "personal",
                "from_library": lib_id,
                # What the template looked like when this copy was made. Later
                # comparisons need both: the template's own fingerprint tells
                # us it moved on, and the copy's tells us whether the owner
                # changed theirs.
                "library_version": library_version(lib),
                # These boards run a dozen or more queries at once, so they
                # opened cold on every visit. Half an hour is current enough
                # for reporting, and any tile can still be marked live.
                "cache_minutes": lib.get("cache_minutes", 30),
                "tabs": tabs,
            }
            dash["library_installed_shape"] = _shape_hash(tabs)
            ok, res = store_dashboard(uid, dash)
            if ok and missing:
                return (True, {"id": res, "missing": missing})
            return (ok, res)
    return False, "Template not found."


# ---------------------------------------------------------------------------
# REPORT CATALOG AND EXECUTION
#
# Report definitions come from the published catalog and are stored locally.
# This script then runs them itself, so report tiles do not depend on any other
# script being installed.
#
# That means SQL arrives from outside, which is a real trust boundary. It is
# handled three ways: nothing executes until an administrator installs it, only
# a single read-only statement is allowed through, and the stored copy is what
# runs, so a change upstream cannot alter a dashboard mid-session.
# ---------------------------------------------------------------------------

# Anything that writes, changes schema, or reaches outside the query. Checked
# as whole words so a column named "update_dt" or "executed" is not rejected.
# Matched as WHOLE words. Without the trailing boundary, "create" rejects any
# report selecting a Created column, which is 39 of the 152 real ones.
_SQL_FORBIDDEN = [
    "insert", "update", "delete", "drop", "truncate", "alter", "create",
    "merge", "grant", "revoke", "exec", "execute",
    "openrowset", "opendatasource", "bulk", "shutdown", "reconfigure",
    "waitfor", "backup", "restore",
    # SELECT ... INTO writes a table. No report in the catalog uses it, so
    # forbidding it costs nothing and closes an obvious hole.
    "into",
]
# Prefixes rather than words: a stored-procedure name has no word boundary
# where the prefix ends.
_SQL_FORBIDDEN_PREFIX = ["sp_", "xp_"]


def sql_is_read_only(sql):
    """(ok, reason). A catalog report must be one read-only statement.

    TouchPoint's own SQL access is read-only, so this is defence in depth
    rather than the only thing standing between a bad catalog and the data.
    It also catches an honest mistake, which is the likelier case.
    """
    if not sql or not sql.strip():
        return (False, "empty")
    t = sql.strip()
    # Strip comments before deciding anything: "-- delete" is not a delete, and
    # more importantly a real statement could hide behind one.
    t = re.sub(r"--[^\n]*", " ", t)
    t = re.sub(r"/\*.*?\*/", " ", t, flags=re.S)
    # Blank out string LITERALS before looking for keywords. 'Delete' is a
    # TouchPoint role name, and a report listing it was refused as though it
    # deleted something. Replaced rather than removed, so the surrounding
    # syntax still reads the same to the checks below.
    t = re.sub(r"'(?:[^']|'')*'", "''", t)
    low = t.lower()
    head = low.lstrip("( \t\r\n")
    if not (head.startswith("select") or head.startswith("with")):
        return (False, "must begin with SELECT or WITH")
    for w in _SQL_FORBIDDEN:
        if re.search(r"\b%s\b" % re.escape(w), low):
            return (False, "contains '%s'" % w)
    for w in _SQL_FORBIDDEN_PREFIX:
        if re.search(r"\b%s" % re.escape(w), low):
            return (False, "contains '%s'" % w)
    # Statement separators: a trailing semicolon is fine, one in the middle is
    # a second statement.
    if ";" in t.rstrip().rstrip(";"):
        return (False, "more than one statement")
    return (True, "")


def load_report_catalog():
    d = load_content_json(CONTENT_REPORTS, {"reports": []})
    if not isinstance(d, dict):
        d = {"reports": []}
    d.setdefault("reports", [])
    return d


def catalog_report(rid):
    for r in load_report_catalog().get("reports", []):
        if str(r.get("id", "")) == str(rid):
            return r
    return None


def ensure_catalog_reports(ids):
    """Install any of these reports the catalog can supply. Returns the set of
    ids the catalog can now serve.

    A template that names a report nobody can supply still installs; that tile
    falls back to Enterprise Reporting, which is the only other source.
    """
    store = load_report_catalog()
    have = {}
    for i, r in enumerate(store.get("reports", [])):
        have[str(r.get("id", ""))] = i
    served = set(k for k in have if k in set(str(x) for x in ids))
    missing = [str(x) for x in ids if str(x) not in have]
    if not missing:
        return served

    items, err = fetch_published_catalog()
    if err:
        return served
    reports = (items.get("reports", []) if isinstance(items, dict)
               else (items or []))
    by_id = dict((str(r.get("id", "")), r) for r in reports)
    added = False
    for rid in missing:
        src = by_id.get(rid)
        if not src:
            continue
        # Checked here as well as at publish: what was safe upstream is not
        # necessarily what arrived.
        ok, why = sql_is_read_only(re.sub(r"\{[a-z_0-9]+\}", "",
                                         src.get("sql_template", "") or ""))
        if not ok:
            continue
        entry = dict(src)
        if rid in have:
            store["reports"][have[rid]] = entry
        else:
            store.setdefault("reports", []).append(entry)
        served.add(rid)
        added = True
    if added:
        save_content_json(CONTENT_REPORTS, store)
    return served


def _catalog_parse(raw):
    if not raw:
        return None, "empty response"
    t = str(raw).strip()
    low = t[:300].lower()
    if "just a moment" in low or "cf-mitigated" in low or "challenge" in low:
        return None, ("Cloudflare challenge intercepted (first 120: "
                      + t[:120].replace("\n", " ") + ")")
    if not t.startswith("{") and not t.startswith("["):
        return None, "non-JSON (first 120: " + t[:120].replace("\n", " ") + ")"
    try:
        return json.loads(t), None
    except Exception as e:
        return None, "JSON parse failed: " + str(e)


_OPS_CACHE = [None]


def fetch_ops_catalog():
    """(items, error). Automatic checks only, since a manual one has no SQL."""
    if _OPS_CACHE[0] is not None:
        return _OPS_CACHE[0], ""
    hdrs = dc_headers()
    last = ""
    for url in (DC_OPS_WORKER, DC_OPS_PUBLIC):
        try:
            resp = model.RestGet(url, hdrs)
        except Exception as e:
            last = "RestGet failed: " + str(e)
            continue
        try:
            data = json.loads(str(resp or ""))
        except Exception as e:
            last = "Could not parse the ops catalog: " + repr(e)
            continue
        out = []
        for it in (data.get("items") or []):
            if str(it.get("type", "")) != "auto":
                continue
            sql = str(it.get("sql", "") or "").strip()
            if not sql:
                continue
            # Marketplace SQL is community-submitted. It is checked here, on
            # the way in, rather than trusted because it was approved upstream.
            ok, why = sql_is_read_only(sql)
            if not ok:
                continue
            out.append({
                "id": str(it.get("name", "")),
                "name": str(it.get("name", "")),
                "cat": str(it.get("cat", "") or "General"),
                "group": str(it.get("defaultGroup", "") or ""),
                "freq": str(it.get("freq", "") or ""),
                "note": str(it.get("inst", "") or ""),
                "steps": [str(x) for x in (it.get("steps") or [])],
                "threshold": it.get("th", 0),
                "sql": sql,
            })
        _OPS_CACHE[0] = out
        return out, ""
    return [], (last or "Could not reach the ops catalog.")


def ops_check(cid):
    items, err = fetch_ops_catalog()
    for it in items:
        if it["id"] == cid:
            return it
    return None


def run_ops_check(cid):
    """Execute one marketplace check and return its rows."""
    it = ops_check(cid)
    if not it:
        return {"success": False,
                "error": "That check is not in the marketplace catalog."}
    sql = it["sql"]
    ok, why = sql_is_read_only(sql)
    if not ok:
        return {"success": False, "error": "Refused to run this check: " + why}
    try:
        rows = q.QuerySql(sql)
    except Exception as e:
        return {"success": False, "error": "Check failed: " + str(e)}
    cols, data, truncated = [], [], False
    for r in rows:
        if not cols:
            cols = row_columns(r, sql)
        row = {}
        for c in cols:
            try:
                v = getattr(r, c, None)
            except Exception:
                v = None
            row[c] = "" if v is None else (
                v if isinstance(v, (int, long, float)) else str(v))
        data.append(row)
        if len(data) >= 2000:
            truncated = True
            break
    return {"success": True, "columns": cols, "rows": data,
            "row_count": len(data), "truncated": truncated,
            "name": it["name"], "note": it["note"], "steps": it["steps"],
            "threshold": it["threshold"],
            # A check is not a report: no date range, no program, no people
            # scope. Said plainly so the bar does not claim otherwise.
            "dated": False, "scopeable": False, "has_program": False}


def fetch_published_catalog():
    """(catalog, error). Browser-posted catalog wins; then workers.dev; then
    the public domain.

    Returns the WHOLE document, not just its reports. It used to hand back
    data["reports"] alone, which silently threw the published dashboards away:
    the catalog carries 16 of them and the Library reported "0 dashboards
    published" for as long as that was true. Every caller already copes with a
    dict, so widening it fixes them all.
    """
    raw = get_param("catalog_json", "")
    if raw:
        data, err = _catalog_parse(raw)
        if data is not None:
            return (data if isinstance(data, dict)
                    else {"reports": data or [], "dashboards": []}, "")
    # No User-Agent here on purpose. model.RestGet is RestSharp, which sets its
    # own and discards the one passed in: the wire shows RestSharp/106.15.0.0
    # whatever this says. Other headers DO come through (RestGet calls
    # AddHeader for each), which is why dc_headers works at all.
    hdrs = dc_headers()
    last = "no attempt made"
    for url in (DC_CATALOG_WORKER, DC_CATALOG_PUBLIC):
        try:
            resp = model.RestGet(url, hdrs)
        except Exception as e:
            last = "RestGet failed: " + str(e)
            continue
        data, err = _catalog_parse(resp)
        if data is not None:
            return (data if isinstance(data, dict)
                    else {"reports": data or [], "dashboards": []}, "")
        last = err
    return ({"reports": [], "dashboards": []}, last)


# The church's fiscal year, as a SQL expression for the first day of the one
# we are currently in. Everything fiscal is derived from this so a church on a
# July or January year gets the same options without a second code path.
def _fy_start(fm):
    return ("CASE WHEN MONTH(GETDATE()) >= %d "
            "THEN DATEFROMPARTS(YEAR(GETDATE()), %d, 1) "
            "ELSE DATEFROMPARTS(YEAR(GETDATE()) - 1, %d, 1) END" % (fm, fm, fm))


def date_range_sql(val):
    """(lower_bound, upper_bound) SQL for a named range. Upper may be None.

    Bounds are half open: >= lower AND < upper.
    """
    if not val:
        return None, None
    try:
        fm = int(load_settings().get("fiscal_year_start_month", 10) or 10)
    except:
        fm = 10
    if fm < 1 or fm > 12:
        fm = 10
    fy = _fy_start(fm)
    jan_this = "DATEFROMPARTS(YEAR(GETDATE()), 1, 1)"
    jan_next = "DATEFROMPARTS(YEAR(GETDATE()) + 1, 1, 1)"
    jan_last = "DATEFROMPARTS(YEAR(GETDATE()) - 1, 1, 1)"

    rolling = {"last_7_days": 7, "last_30_days": 30, "last_60_days": 60,
               "last_90_days": 90, "last_180_days": 180, "last_365_days": 365}
    if val in rolling:
        return "DATEADD(day, -%d, GETDATE())" % rolling[val], None
    months = {"last_3_months": 3, "last_6_months": 6, "last_12_months": 12,
              "last_24_months": 24, "last_36_months": 36, "last_48_months": 48}
    if val in months:
        return "DATEADD(month, -%d, GETDATE())" % months[val], None

    if val == "ytd":                 # kept: saved tiles already use this name
        return jan_this, None
    if val == "this_calendar_year":
        return jan_this, jan_next
    if val == "last_calendar_year":
        return jan_last, jan_this
    # fiscal_ytd is the name Enterprise Reporting uses and the one ten catalog
    # reports carry as their DEFAULT. Nothing here implemented it, so those ran
    # with no date bound at all -- the whole contribution history rather than
    # the year. fytd is accepted as an alias.
    if val in ("fiscal_ytd", "fytd"):
        return "(%s)" % fy, None
    if val == "this_fiscal_year":
        return "(%s)" % fy, "DATEADD(year, 1, %s)" % fy
    if val == "last_fiscal_year":
        return "DATEADD(year, -1, %s)" % fy, "(%s)" % fy
    return None, None


def build_filter_sql(params, values):
    """WHERE fragments from a report's declared parameters.

    Only parameters the DEFINITION declares can produce SQL, and each value is
    quoted or checked numeric. A tile cannot introduce a clause the report did
    not offer.
    """
    out = []
    for p_ in (params or []):
        name = p_.get("name", "")
        col = p_.get("sql_column", "")
        ptype = p_.get("type", "text")
        # Fall back to the parameter's own default. A tile stores only the
        # filters that were set, so without this a report written around
        # 'last_90_days' runs with no date bound at all -- mem_new_members
        # returns 15,903 rows instead of a quarter's worth.
        val = str(values.get(name, "") or "")
        if not val:
            val = str(p_.get("default", "") or "")
        if not col or not val:
            continue
        if not re.match(r"^[A-Za-z0-9_\.\[\]]+$", col):
            continue                       # a column name, not an expression
        if ptype == "daterange":
            lo, hi = date_range_sql(val)
            if lo:
                out.append("%s >= %s" % (col, lo))
                # Half open on purpose. A datetime column compared with
                # BETWEEN against a date drops everything after midnight on
                # the last day, which silently loses a day of data.
                if hi:
                    out.append("%s < %s" % (col, hi))
            elif "|" in val:
                a, b = val.split("|", 1)
                a = re.sub(r"[^0-9\-/]", "", a)
                b = re.sub(r"[^0-9\-/]", "", b)
                if a and b:
                    out.append("%s BETWEEN '%s' AND '%s'" % (col, a, b))
        elif ptype == "multi_select":
            ids = [v.strip() for v in val.split(",") if v.strip()]
            if ids and all(v.isdigit() for v in ids):
                out.append("%s IN (%s)" % (col, ",".join(ids)))
            elif ids:
                out.append("%s IN (%s)" % (col, ",".join(
                    "'" + v.replace("'", "''") + "'" for v in ids)))
        elif ptype == "number":
            try:
                out.append("%s >= %s" % (col, float(val)))
            except:
                pass
        else:
            if val.isdigit():
                out.append("%s = %s" % (col, val))
            else:
                out.append("%s = '%s'" % (col, val.replace("'", "''")))
    return out


# Columns off a QuerySql row. A row is a .NET dynamic object, so __dict__ is
# empty on it -- reading that gave every report zero columns, which rendered as
# a blank tile rather than an error. dir() is what actually enumerates them.
_ROW_NOISE = set(["Count", "Keys", "Values", "Items", "GetType", "ToString",
                  "Equals", "GetHashCode", "ReferenceEquals", "MemberwiseClone",
                  "Finalize", "IsReadOnly", "Item"])


def sql_select_columns(sql):
    """Column names in SELECT order, so a table reads like the report intended.

    dir() comes back alphabetical, which would scramble the column order and,
    worse, hand a chart the wrong label column by default. Everything here is
    depth-aware: with CTEs the first SELECTs belong to the CTE bodies, and an
    outer select list routinely contains subqueries carrying their own FROM,
    so a plain regex picks the wrong clause on exactly the reports that matter.
    """
    try:
        txt = re.sub(r"--[^\n]*", " ", sql or "")
        txt = re.sub(r"/\*.*?\*/", " ", txt, flags=re.DOTALL)
        txt = re.sub(r"'[^']*'", "''", txt)

        # Paren depth at every character, so tokens can be judged by nesting.
        depth, depths = 0, []
        for ch in txt:
            if ch == "(":
                depth += 1
                depths.append(depth - 1)   # the "(" itself sits outside
            elif ch == ")":
                depth -= 1
                depths.append(depth)
            else:
                depths.append(depth)

        def token_at(word, frm):
            for m in re.finditer(r"\b" + word + r"\b", txt[frm:], re.IGNORECASE):
                i = frm + m.start()
                if depths[i] == 0:
                    return i, frm + m.end()
            return -1, -1

        # The first SELECT at depth 0 is the result set: CTE bodies are
        # parenthesized, so they never qualify.
        s0, s1 = token_at("SELECT", 0)
        if s0 < 0:
            return []
        f0, _ = token_at("FROM", s1)
        part = txt[s1:f0] if f0 > 0 else txt[s1:]
        part = re.sub(r"^\s*(?:DISTINCT\s+)?(?:TOP\s+\d+\s+)?", " ", part,
                      flags=re.IGNORECASE)

        # Split on commas that are not inside parentheses, so a subquery or a
        # CASE expression in the select list stays one column.
        out, d, buf = [], 0, []
        for ch in part:
            if ch == "(":
                d += 1
            elif ch == ")":
                d -= 1
            if ch == "," and d <= 0:
                out.append("".join(buf))
                buf = []
            else:
                buf.append(ch)
        out.append("".join(buf))

        cols = []
        for chunk in out:
            chunk = chunk.strip()
            if not chunk:
                continue
            a = re.search(r"\bAS\s+\[?(\w+)\]?\s*$", chunk, re.IGNORECASE)
            if a:
                cols.append(a.group(1))
                continue
            w = re.search(r"(\w+)\s*$", chunk)
            # An unaliased CASE ends in END, which is a keyword, not a name.
            # Only ordering is at stake here, so dropping it is safe: the real
            # name still arrives from the row itself.
            if w and w.group(1).upper() not in ("END", "ELSE", "NULL"):
                cols.append(w.group(1))
        return cols
    except:
        return []


def row_columns(r, sql):
    avail = set()
    for attr in dir(r):
        if attr.startswith("_") or attr in _ROW_NOISE:
            continue
        try:
            getattr(r, attr)
            avail.add(attr)
        except:
            pass
    cols = []
    lower = dict((a.lower(), a) for a in avail)
    for name in sql_select_columns(sql):
        actual = lower.get(name.lower())
        if actual and actual not in cols:
            cols.append(actual)
    for a in sorted(avail):
        if a not in cols:
            cols.append(a)
    return cols


# A people scope is injected as `p.PeopleId IN (...)`, so it only works when
# the report actually binds `p` to People. One catalog entry declares itself
# scopeable while aliasing People as something else, which turned a saved
# search into an "Invalid column name 'p'" failure. Checked rather than
# trusted, so a wrong upstream flag cannot produce a broken tile.
# {cy}..{cy9} are this calendar year and the nine before it. Two reports
# alias their year columns as [{cy}] and name the same columns in their display
# block, so the token has to resolve identically in both places or the chart
# looks for a column called "{cy}" and finds nothing.
def resolve_year_tokens(text):
    if not text or "{cy" not in text:
        return text
    year = datetime.datetime.now().year
    for back in range(9, -1, -1):
        text = text.replace("{cy%s}" % (back or ""), str(year - back))
    return text


# Placeholders that stand in for a whole clause or column list. Unlike {cy}
# these are never column names, so they are removed rather than resolved.
def strip_injection_tokens(sql):
    return re.sub(r"\{(?:demographics_\w+|filters|filters_where"
                  r"|bluetoolbar_filter|current_org_filter)\}", " ", sql or "")


def resolve_display(disp):
    """A display block with its year tokens turned into real column names."""
    if not disp:
        return {}
    out = {}
    for k, v in disp.items():
        if isinstance(v, (list, tuple)):
            out[k] = [resolve_year_tokens(x) if isinstance(x, str) else x
                      for x in v]
        elif isinstance(v, str):
            out[k] = resolve_year_tokens(v)
        else:
            out[k] = v
    return out


def accepts_people_scope(rep):
    sql = rep.get("sql_template", "") or ""
    if "{filters}" not in sql and "{filters_where}" not in sql:
        return False
    return bool(re.search(r"\bPeople\s+(?:AS\s+)?p\b", sql, re.IGNORECASE))


# Reports name this parameter either way, so both spellings are accepted
# rather than making the dashboard care which one a report happened to use.
PROGRAM_PARAMS = ("program", "program_id")


def report_caps(rep):
    """What a dashboard may narrow this report by.

    The published catalog states this per report, decided once at publish
    time. Deriving it here as well would mean two implementations that can
    disagree, and the SQL test for people-scope is a regex nobody should own
    twice. Older copies installed before the catalog carried the flags are
    still worked out locally so they keep behaving.
    """
    caps = rep.get("caps")
    if isinstance(caps, dict) and caps:
        return {"date": bool(caps.get("date")),
                "program": bool(caps.get("program")),
                "people": bool(caps.get("people"))}
    return {"date": bool(report_date_param(rep)),
            "program": bool(report_program_param(rep)),
            "people": accepts_people_scope(rep)}


def report_program_param(rep):
    """The name of the report's program filter, if it has one."""
    for prm in (rep.get("parameters") or []):
        if prm.get("name") in PROGRAM_PARAMS:
            return prm.get("name")
    return None


def list_programs():
    """Programs, for the dashboard-wide area filter."""
    out = []
    try:
        for r in q.QuerySql("SELECT Id, Name FROM dbo.Program WITH (NOLOCK) "
                            "WHERE ISNULL(Name,'') <> '' ORDER BY Name") or []:
            out.append({"id": int(r.Id), "name": str(r.Name)})
    except Exception:
        pass
    return out


def report_date_param(rep):
    """The name of the report's date-range parameter, if it has one."""
    for prm in (rep.get("parameters") or []):
        if prm.get("type") == "daterange":
            return prm.get("name") or "date_range"
    return None


def run_catalog_report(rid, values, people_ids=None, serving=None,
                       global_range=None, global_query=None,
                       global_program=None):
    rep = catalog_report(rid)
    if not rep:
        return {"success": False,
                "error": "Report '%s' is not installed from the catalog." % rid}
    # Before any SQL is built. A tile hidden in the picker is still reachable
    # by posting run_catalog directly, and a shared dashboard can carry a
    # finance tile to someone who should not see it.
    cat = rep.get("category", "")
    if not can_view_category(cat):
        # Flagged, not just worded: the tile renders this as a quiet
        # placeholder rather than a red error, so a shared dashboard does not
        # look broken to someone who simply is not on the finance team.
        return {"success": False, "restricted": True,
                "error": "You do not have access to this data."}
    sql = rep.get("sql_template", "") or ""

    # A dashboard-wide date range answers "over what period?" once, instead of
    # per tile. It only fills a gap: a tile that set its own range keeps it, and
    # a report with no date parameter is left alone rather than being silently
    # reinterpreted.
    # A dashboard-wide search answers "which people?" once. Like the date
    # range it only fills a gap: a tile with its own search keeps it, and a
    # report that cannot be limited to people is left alone rather than being
    # refused outright, which is what passing ids to it would do.
    scopeable = accepts_people_scope(rep)
    scoped = ""
    if people_ids is None and global_query and scopeable:
        got, note, err = _scope_ids(global_query)
        if err:
            return {"success": False, "error": err}
        people_ids = got
        scoped = global_query

    values = dict(values or {})
    # Attendance belongs to a program, and most reports here already declare a
    # program filter. That is the right instrument for "only my area": a
    # people search cannot narrow a report that counts head counts per meeting,
    # but a program filter can, because the meeting knows its program.
    pp = report_program_param(rep)
    has_program = bool(pp)
    program_scope = ""
    if global_program and pp and not str(values.get(pp, "")).strip():
        values[pp] = global_program
        program_scope = global_program

    # Recorded, not inferred: the tile shows whether the range was applied
    # based on what happened here, so "this report has no dates" can never be
    # confused with "the range is broken".
    date_scope = ""
    dated = bool(report_date_param(rep))
    if global_range:
        dp = report_date_param(rep)
        if dp and not str(values.get(dp, "")).strip():
            values[dp] = global_range
            date_scope = global_range

    where = build_filter_sql(rep.get("parameters", []), values or {})
    if people_ids:
        if not accepts_people_scope(rep):
            return {"success": False,
                    "error": "This report cannot be limited to a search: it "
                             "does not select people through the expected "
                             "alias. Remove the search from this tile."}
        ids = [str(int(x)) for x in people_ids][:5000]
        if ids:
            where.append("p.PeopleId IN (%s)" % ",".join(ids))

    sql = sql.replace("{filters}", ("AND " + " AND ".join(where)) if where else "")
    sql = sql.replace("{filters_where}",
                      ("WHERE " + " AND ".join(where)) if where else "")
    sql = sql.replace("{bluetoolbar_filter}", "")
    sql = sql.replace("{current_org_filter}", "")

    # Church-specific values the definition declared it needs.
    st = load_settings()
    serv = ",".join(str(int(x)) for x in (serving or [])
                    if str(x).strip().isdigit()) or \
        str(st.get("serving_member_types", "140,310,320,710"))
    sql = sql.replace("{serving_types}", serv)
    sql = sql.replace("{leader_types}",
                      str(st.get("leader_member_types", "140,310,320")))
    sql = sql.replace("{active_status_types}",
                      str(st.get("active_status_types", "10,20,30")))
    # How long a background check stays valid is church policy, not a
    # constant. It was hardcoded at 730 days inside the report SQL.
    sql = sql.replace("{bg_check_days}",
                      str(int(st.get("bg_check_days", 730) or 730)))

    # Optional demographic columns. Empty is the "none added" case upstream, so
    # the report runs with its own columns only.
    for d in ("select", "join", "groupby", "passthrough"):
        sql = sql.replace("{demographics_%s}" % d, "")

    # {cy}..{cy9}: this calendar year and the nine before it. Upstream bakes
    # these in when it loads its definitions, so they survive into the exported
    # template and have to be resolved here.
    sql = resolve_year_tokens(sql)

    # Default on: matching the church's own weekly report is the less
    # surprising answer when the two are compared side by side.
    approvals, approvals_set = bg_approval_list()
    sql = sql.replace("{bg_approval_ok}",
                      ",".join("'" + a.replace("'", "''") + "'"
                               for a in approvals))

    sql = sql.replace("{waag_filter}",
                      "" if str(st.get("attendance_scope", "waag")) == "all"
                      else WAAG_SQL)

    fm = str(st.get("fiscal_year_start_month", 10))
    sql = sql.replace("{fiscal_month}", fm)
    sql = sql.replace("{year_start}",
                      fm if (values or {}).get("year_type") == "fiscal" else "1")

    # {range_start} / {range_end}: the chosen window as actual bounds rather
    # than a WHERE fragment. Reports that compare periods against each other
    # have to slice the range themselves, which {filters} cannot express.
    if "{range_start}" in sql or "{range_end}" in sql:
        dp = report_date_param(rep)
        chosen = str((values or {}).get(dp, "")).strip() if dp else ""
        if not chosen and dp:
            for prm in (rep.get("parameters") or []):
                if prm.get("name") == dp:
                    chosen = str(prm.get("default", "") or "")
        lo, hi = date_range_sql(chosen)
        if not lo:
            return {"success": False,
                    "error": "This report needs a date range and none is set."}
        sql = sql.replace("{range_start}", "(" + lo + ")")
        sql = sql.replace("{range_end}", "(" + (hi or "GETDATE()") + ")")

    # Anything still unresolved would reach SQL Server as a syntax error, and
    # the tile would just sit there empty. Name it instead.
    left = sorted(set(re.findall(r"\{[a-z_0-9]+\}", sql)))
    if left:
        return {"success": False,
                "error": "This report uses %s, which this script does not "
                         "know how to fill in." % ", ".join(left)}

    ok, why = sql_is_read_only(sql)
    if not ok:
        return {"success": False,
                "error": "Refused to run this report: %s. Catalog reports must "
                         "be a single read-only SELECT." % why}
    try:
        rows = q.QuerySql(sql)
    except Exception as e:
        return {"success": False, "error": "Query failed: " + str(e)}

    cols, data = [], []
    truncated = False
    # 1,000 was too low for a work list: the background-check report alone is
    # 1,295 people, so a fifth of it was never read. Raised, and settable per
    # church because the ceiling that matters is the browser's, not ours.
    try:
        row_cap = int(load_settings().get("max_rows", 1500) or 1500)
    except:
        row_cap = 1500
    # 5,000 was the old default, and it is why a people-list tile weighed
    # 666 KB: five thousand rows fetched, serialised and held in order to draw
    # two hundred. That put it over the cache ceiling, so those tiles re-ran
    # every single time -- the slow path AND no caching. 1,500 still fills the
    # table's paging comfortably and the truncation note says when more exist.
    # A church that genuinely wants more can raise max_rows.
    if row_cap < 100:
        row_cap = 100
    if row_cap > 20000:
        row_cap = 20000
    for r in rows:
        if not cols:
            cols = row_columns(r, sql)
        row = {}
        for c in cols:
            try:
                v = getattr(r, c, None)
            except:
                v = None
            if v is None:
                row[c] = ""
            else:
                try:
                    row[c] = v if isinstance(v, (int, long, float)) else str(v)
                except:
                    row[c] = str(v)
        data.append(row)
        if len(data) >= row_cap:
            # Stop, but remember that we did. A tile counting these rows would
            # otherwise print exactly 1,000 and read as a measurement: the
            # background-check list is 1,295 people and reported "1,000".
            truncated = True
            break
    return {"success": True, "columns": cols, "rows": data,
            "row_count": len(data), "truncated": truncated,
            "display": resolve_display(rep.get("display", {}) or {}),
            "date_scope": date_scope, "dated": dated,
            "people_scope": scoped, "scopeable": scopeable,
            "needs_setting": ("bg_approval_ok"
                              if ("{bg_approval_ok}" in (rep.get("sql_template") or "")
                                  and not approvals_set) else ""),
            "program_scope": program_scope, "has_program": has_program,
            "name": rep.get("name", rid)}


# ---------------------------------------------------------------------------
# TOUCHPOINT WIDGETS
# TouchPoint's own home-page widgets render at /HomeWidgets/Embed/{id}. They
# already carry HTML + Python + SQL, a cache policy, and per-widget roles, so a
# widget tile costs us nothing but the fetch.
# ---------------------------------------------------------------------------
def list_touchpoint_widgets():
    """Every widget this person could embed.

    Deliberately NOT filtered to Enabled=1. That column only controls whether
    TouchPoint shows a widget on its own home page; the Embed endpoint loads a
    widget by id and checks roles alone. Filtering on it hid 30 of 44 widgets
    here, including most of the ones worth putting on a dashboard.
    """
    sql = """
        SELECT TOP 200 w.Id, w.Name, w.Description, w.CacheHours, w.Enabled,
               ISNULL(STUFF((SELECT ',' + ro.RoleName
                             FROM DashboardWidgetRoles dr
                             JOIN Roles ro ON ro.RoleId = dr.RoleId
                             WHERE dr.WidgetId = w.Id
                             FOR XML PATH('')), 1, 1, ''), '') AS RoleNames
        FROM DashboardWidgets w
        ORDER BY w.Enabled DESC, w.[Order], w.Id
    """
    out = []
    skipped = {}
    try:
        rows = q.QuerySql(sql)
    except Exception:
        return out, skipped
    for r in rows:
        names = [x for x in str(getattr(r, "RoleNames", "") or "").split(",") if x]
        # TouchPoint's Embed action requires a NON-EMPTY role intersection, so a
        # widget with no roles can never be embedded even though the home page
        # happily shows it. Offering one would produce a tile that always reads
        # "Not authorized", so leave it out.
        if not names:
            skipped["no_roles"] = skipped.get("no_roles", 0) + 1
            continue
        allowed = False
        for n in names:
            try:
                if model.UserIsInRole(n):
                    allowed = True
                    break
            except Exception:
                pass
        if not allowed:
            skipped["not_mine"] = skipped.get("not_mine", 0) + 1
            continue
        # DashboardWidgetRoles holds duplicates (one widget here has 67 rows
        # for 24 distinct roles), and a widget open to everything would print a
        # 200-character line in the picker. Dedupe and summarize.
        uniq = []
        for x in names:
            if x not in uniq:
                uniq.append(x)
        shown = ", ".join(uniq[:4])
        if len(uniq) > 4:
            shown += " +%d more" % (len(uniq) - 4)
        out.append({
            "id": int(r.Id),
            "name": str(getattr(r, "Name", "") or ""),
            "description": str(getattr(r, "Description", "") or ""),
            "cache_hours": int(getattr(r, "CacheHours", 0) or 0),
            # Enabled governs TouchPoint's HOME PAGE only. Embed does not check
            # it, so a widget switched off there still works as a tile -- and
            # the analytical ones (giving trends, attendance year over year)
            # are mostly switched off. Reported, not filtered on.
            "enabled": (1 if int(getattr(r, "Enabled", 0) or 0) else 0),
            "roles": shown,
        })
    return out, skipped


# ---------------------------------------------------------------------------
# CUSTOM TILES
# Build a tile from a Search Builder search plus a choice of what to measure.
#
# EVERY fragment of SQL below is a constant in this file. The browser only ever
# sends KEYS into these tables, never SQL, never column names. A tile is
# therefore incapable of expressing a query this file does not already contain,
# which is the only safe way to let a UI assemble SQL.
# ---------------------------------------------------------------------------

# Ceiling on ids pulled from a saved search, matching Enterprise Reporting.
CUSTOM_MAX_IDS = 5000

_AGE_BIN = ("CASE WHEN p.Age IS NULL THEN 'Unknown'"
            " WHEN p.Age < 13 THEN '0-12'"
            " WHEN p.Age < 18 THEN '13-17'"
            " WHEN p.Age < 25 THEN '18-24'"
            " WHEN p.Age < 35 THEN '25-34'"
            " WHEN p.Age < 45 THEN '35-44'"
            " WHEN p.Age < 55 THEN '45-54'"
            " WHEN p.Age < 65 THEN '55-64'"
            " WHEN p.Age < 75 THEN '65-74'"
            " ELSE '75+' END")

_AGE_SORT = ("CASE WHEN p.Age IS NULL THEN 999 ELSE p.Age / 10 END")

CUSTOM_FILTERS = {
    "campus": {
        "label": "Campus",
        "col": "p.CampusId",
        "sql": "SELECT Id AS v, Description AS n FROM lookup.Campus "
               "WHERE Id > 0 ORDER BY Description",
    },
    "member_status": {
        "label": "Member status",
        "col": "p.MemberStatusId",
        "sql": "SELECT Id AS v, Description AS n FROM lookup.MemberStatus "
               "ORDER BY Id",
    },
}


def custom_filter_sql(spec):
    """WHERE fragments for the filters a built tile has set.

    Values are forced to integers rather than quoted: these are lookup ids,
    and an id is the only thing that can legitimately arrive here.
    """
    out = []
    for key in sorted(CUSTOM_FILTERS.keys()):
        f = CUSTOM_FILTERS[key]
        raw = str((spec or {}).get("f_" + key, "") or "").strip()
        if not raw:
            continue
        ids = [x.strip() for x in raw.split(",")
               if x.strip().lstrip("-").isdigit()]
        if not ids:
            continue
        out.append("%s IN (%s)" % (f["col"], ",".join(str(int(x)) for x in ids)))
    return out


CUSTOM_DOMAINS = {
    "demographics": {
        "label": "Demographics",
        "help": "Who these people are.",
        "dates": False,
        "from": ("FROM dbo.People p WITH (NOLOCK)"),
        "where": "p.IsDeceased = 0 AND p.ArchivedFlag = 0",
        "person_col": "p.PeopleId",
        "measures": {
            "count_people": {"label": "Number of people",
                             "expr": "COUNT(DISTINCT p.PeopleId)", "money": False},
        },
        "dimensions": {
            "none":        {"label": "One total", "expr": None, "join": ""},
            "age_bin":     {"label": "Age band", "expr": _AGE_BIN, "join": "",
                            "order": _AGE_SORT},
            "gender":      {"label": "Gender",
                            "expr": "ISNULL(g.Description,'Unknown')",
                            "join": "LEFT JOIN lookup.Gender g ON g.Id = p.GenderId"},
            "member_status": {"label": "Member status",
                              "expr": "ISNULL(ms.Description,'Unknown')",
                              "join": "LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId"},
            "campus":      {"label": "Campus",
                            "expr": "ISNULL(cp.Description,'None')",
                            "join": "LEFT JOIN lookup.Campus cp ON cp.Id = p.CampusId"},
            "marital":     {"label": "Marital status",
                            "expr": "ISNULL(mar.Description,'Unknown')",
                            "join": "LEFT JOIN lookup.MaritalStatus mar ON mar.Id = p.MaritalStatusId"},
            "family_pos":  {"label": "Position in family",
                            "expr": "ISNULL(fp.Description,'Unknown')",
                            "join": "LEFT JOIN lookup.FamilyPosition fp ON fp.Id = p.PositionInFamilyId"},
        },
    },

    "involvements": {
        "label": "Involvements",
        "help": "Who is enrolled in what, as the rosters stand now.",
        "dates": True,
        "from": ("FROM dbo.OrganizationMembers om WITH (NOLOCK)"
                 " JOIN dbo.Organizations o ON o.OrganizationId = om.OrganizationId"
                 " JOIN dbo.People p ON p.PeopleId = om.PeopleId"),
        # Active enrolments in active involvements. A dropped member still has
        # a row, so without InactiveDate this counts everyone who ever joined.
        "where": ("om.InactiveDate IS NULL AND o.OrganizationStatusId = 30"
                  " AND p.IsDeceased = 0"),
        "person_col": "om.PeopleId",
        "date_col": "om.EnrollmentDate",
        "measures": {
            "count_people": {"label": "Number of people",
                             "expr": "COUNT(DISTINCT om.PeopleId)", "money": False},
            "count_enroll": {"label": "Number of enrolments",
                             "expr": "COUNT(*)", "money": False},
            "count_orgs":   {"label": "Number of involvements",
                             "expr": "COUNT(DISTINCT om.OrganizationId)",
                             "money": False},
            "avg_per_person": {"label": "Average involvements per person",
                               "expr": ("CAST(COUNT(*) * 1.0 /"
                                        " NULLIF(COUNT(DISTINCT om.PeopleId),0)"
                                        " AS DECIMAL(10,1))"), "money": False},
        },
        "dimensions": {
            "none":     {"label": "One total", "expr": None, "join": ""},
            "month":    {"label": "Month joined",
                         "expr": "CONVERT(varchar(7), om.EnrollmentDate, 120)",
                         "join": ""},
            "program":  {"label": "Program",
                         "expr": "ISNULL(ipr.Name,'Unassigned')",
                         "join": ("LEFT JOIN dbo.Division idv ON idv.Id = o.DivisionId"
                                  " LEFT JOIN dbo.Program ipr ON ipr.Id = idv.ProgId")},
            "division": {"label": "Division",
                         "expr": "ISNULL(idv2.Name,'Unassigned')",
                         "join": "LEFT JOIN dbo.Division idv2 ON idv2.Id = o.DivisionId"},
            "involvement": {"label": "Involvement",
                            "expr": "o.OrganizationName", "join": ""},
            "member_type": {"label": "Member type",
                            "expr": "ISNULL(imt.Description,'Unknown')",
                            "join": "LEFT JOIN lookup.MemberType imt ON imt.Id = om.MemberTypeId"},
            "age_bin":  {"label": "Age band", "expr": _AGE_BIN, "join": "",
                         "order": _AGE_SORT},
        },
    },

    "tasks": {
        "label": "Tasks and notes",
        "help": "Contact work recorded on people.",
        "dates": True,
        # dbo.Task is empty and legacy; TaskNote holds both, and IsNote tells
        # them apart. Statuses come from DbUtil.Codes: 1 complete, 6 archived.
        "from": ("FROM dbo.TaskNote tn WITH (NOLOCK)"
                 " JOIN dbo.People p ON p.PeopleId = tn.AboutPersonId"),
        "where": "p.IsDeceased = 0",
        "person_col": "tn.AboutPersonId",
        "date_col": "tn.CreatedDate",
        "measures": {
            "count_items":  {"label": "Number of tasks and notes",
                             "expr": "COUNT(*)", "money": False},
            "count_people": {"label": "People contacted",
                             "expr": "COUNT(DISTINCT tn.AboutPersonId)",
                             "money": False},
            "count_open":   {"label": "Still open",
                             "expr": ("SUM(CASE WHEN tn.IsNote = 0"
                                      " AND ISNULL(tn.StatusId,0) NOT IN (1,6)"
                                      " THEN 1 ELSE 0 END)"), "money": False},
            "count_done":   {"label": "Completed",
                             "expr": ("SUM(CASE WHEN tn.StatusId = 1"
                                      " THEN 1 ELSE 0 END)"), "money": False},
        },
        "dimensions": {
            "none":   {"label": "One total", "expr": None, "join": ""},
            "month":  {"label": "Month",
                       "expr": "CONVERT(varchar(7), tn.CreatedDate, 120)",
                       "join": ""},
            "kind":   {"label": "Task or note",
                       "expr": ("CASE WHEN tn.IsNote = 1 THEN 'Note'"
                                " ELSE 'Task' END"), "join": ""},
            "owner":  {"label": "Assigned to",
                       "expr": "ISNULL(own.Name2,'Unassigned')",
                       "join": "LEFT JOIN dbo.People own ON own.PeopleId = tn.OwnerId"},
            "age_bin": {"label": "Age band", "expr": _AGE_BIN, "join": "",
                        "order": _AGE_SORT},
        },
    },

    "attendance": {
        "label": "Attendance",
        "help": "Meetings these people actually attended.",
        "dates": True,
        # AttendanceFlag = 1 is the difference between "was marked present" and
        # "appears on a roster". Anything else overcounts badly.
        "from": ("FROM dbo.Attend a WITH (NOLOCK)"
                 " JOIN dbo.Meetings m ON m.MeetingId = a.MeetingId"
                 " JOIN dbo.People p ON p.PeopleId = a.PeopleId"),
        "where": ("a.AttendanceFlag = 1 AND m.DidNotMeet = 0"
                  " AND p.IsDeceased = 0"),
        "person_col": "a.PeopleId",
        "date_col": "m.MeetingDate",
        "measures": {
            "count_attends": {"label": "Times attended",
                              "expr": "COUNT(*)", "money": False},
            "count_people":  {"label": "Distinct people",
                              "expr": "COUNT(DISTINCT a.PeopleId)", "money": False},
            "avg_per_person": {"label": "Average visits per person",
                               "expr": ("CAST(COUNT(*) * 1.0 /"
                                        " NULLIF(COUNT(DISTINCT a.PeopleId),0)"
                                        " AS DECIMAL(10,1))"), "money": False},
        },
        "dimensions": {
            "none":    {"label": "One total", "expr": None, "join": ""},
            "month":   {"label": "Month",
                        "expr": "CONVERT(varchar(7), m.MeetingDate, 120)", "join": ""},
            "program": {"label": "Program",
                        "expr": "ISNULL(pr.Name,'Unassigned')",
                        "join": ("LEFT JOIN dbo.Organizations o ON o.OrganizationId = a.OrganizationId"
                                 " LEFT JOIN dbo.Division dv ON dv.Id = o.DivisionId"
                                 " LEFT JOIN dbo.Program pr ON pr.Id = dv.ProgId")},
            "division": {"label": "Division",
                         "expr": "ISNULL(dv2.Name,'Unassigned')",
                         "join": ("LEFT JOIN dbo.Organizations o2 ON o2.OrganizationId = a.OrganizationId"
                                  " LEFT JOIN dbo.Division dv2 ON dv2.Id = o2.DivisionId")},
            "involvement": {"label": "Involvement",
                            "expr": "ISNULL(o3.OrganizationName,'Unknown')",
                            "join": "LEFT JOIN dbo.Organizations o3 ON o3.OrganizationId = a.OrganizationId"},
            "age_bin": {"label": "Age band", "expr": _AGE_BIN, "join": "",
                        "order": _AGE_SORT},
        },
    },

    "giving": {
        "label": "Giving",
        "help": "Contributions from these people.",
        "dates": True,
        # Types 6/7/8/99 are returned, reversed, pledged and event fees. None
        # of them are giving, and including them silently inflates every total.
        "from": ("FROM dbo.Contribution c WITH (NOLOCK)"
                 " JOIN dbo.People p ON p.PeopleId = c.PeopleId"),
        "where": ("c.ContributionTypeId NOT IN (6,7,8,99)"
                  " AND c.ContributionStatusId = 0"),
        "person_col": "c.PeopleId",
        "date_col": "c.ContributionDate",
        "measures": {
            "sum_amount":   {"label": "Total given",
                             "expr": "SUM(c.ContributionAmount)", "money": True},
            "avg_amount":   {"label": "Average gift",
                             "expr": "AVG(c.ContributionAmount)", "money": True},
            "count_gifts":  {"label": "Number of gifts",
                             "expr": "COUNT(*)", "money": False},
            "count_donors": {"label": "Distinct givers",
                             "expr": "COUNT(DISTINCT c.PeopleId)", "money": False},
            "avg_per_donor": {"label": "Average per giver",
                              "expr": ("CAST(SUM(c.ContributionAmount) /"
                                       " NULLIF(COUNT(DISTINCT c.PeopleId),0)"
                                       " AS DECIMAL(18,2))"), "money": True},
        },
        "dimensions": {
            "none":  {"label": "One total", "expr": None, "join": ""},
            "month": {"label": "Month",
                      "expr": "CONVERT(varchar(7), c.ContributionDate, 120)", "join": ""},
            "fund":  {"label": "Fund",
                      "expr": "ISNULL(cf.FundName,'Unknown')",
                      "join": "LEFT JOIN dbo.ContributionFund cf ON cf.FundId = c.FundId"},
            "age_bin": {"label": "Age band", "expr": _AGE_BIN, "join": "",
                        "order": _AGE_SORT},
            "campus": {"label": "Campus",
                       "expr": "ISNULL(cp2.Description,'None')",
                       "join": "LEFT JOIN lookup.Campus cp2 ON cp2.Id = p.CampusId"},
        },
    },
}


def custom_filter_meta():
    """Filter definitions plus their options, for the tile builder."""
    out = []
    for key in sorted(CUSTOM_FILTERS.keys()):
        f = CUSTOM_FILTERS[key]
        opts = []
        try:
            for r in q.QuerySql(f["sql"]):
                opts.append({"v": str(getattr(r, "v", "")),
                             "n": str(getattr(r, "n", "") or "")})
        except Exception:
            opts = []
        # A filter with nothing to choose from is not a choice. This church
        # has no campuses at all, and offering an empty Campus box is the
        # same "looks broken" problem as a bar control that does nothing.
        if opts:
            out.append({"key": key, "label": f["label"], "options": opts})
    return out


def custom_meta():
    """The catalog, shaped for the browser. Keys only: no SQL crosses the wire."""
    out = []
    for key in ("demographics", "attendance", "giving"):
        d = CUSTOM_DOMAINS[key]
        out.append({
            "key": key,
            "label": d["label"],
            "help": d["help"],
            "dates": d["dates"],
            "measures": [{"key": k, "label": v["label"]}
                         for k, v in sorted(d["measures"].items(),
                                            key=lambda kv: kv[1]["label"])],
            "dimensions": [{"key": k, "label": v["label"]}
                           for k, v in sorted(d["dimensions"].items(),
                                              key=lambda kv: (kv[0] != "none",
                                                              kv[1]["label"]))],
        })
    return out


def list_saved_searches():
    """Named Search Builder searches this person may use.

    Search Builder autosaves work in progress into the same table under the
    name 'Draft', which is scratch state and not a search anyone chose to
    keep. Listing those buries the real ones.

    Two things this deliberately does NOT do:
      - cap the list. It used to take the first 400, and this database has 557
        visible to one user, so 157 were dropped with nothing said.
      - let a public search hide one of your own with the same name. The
        de-duplication is by name, so it prefers yours.
    """
    uid = current_user_id()
    sql = """
        WITH mine AS (
            SELECT Username FROM dbo.Users WITH (NOLOCK)
            WHERE PeopleId = {uid} AND ISNULL(Username,'') <> ''
        ),
        vis AS (
            SELECT qq.name AS Nm, ISNULL(qq.owner,'') AS Owner,
                   CASE WHEN qq.owner IN (SELECT Username FROM mine)
                        THEN 1 ELSE 0 END AS IsMine,
                   qq.lastRun, qq.created
            FROM dbo.Query qq WITH (NOLOCK)
            WHERE ISNULL(qq.name,'') <> ''
              AND qq.name NOT IN ('Draft', 'OrgFilter')
              AND (qq.ispublic = 1 OR qq.owner IN (SELECT Username FROM mine))
        ),
        ranked AS (
            SELECT Nm, Owner, IsMine,
                   ROW_NUMBER() OVER (PARTITION BY Nm
                        ORDER BY IsMine DESC, lastRun DESC, created DESC) AS rn
            FROM vis
        )
        SELECT Nm AS Name, Owner, IsMine FROM ranked WHERE rn = 1 ORDER BY Nm
    """.format(uid=int(uid or 0))
    out = []
    try:
        for r in q.QuerySql(sql):
            out.append({"name": str(getattr(r, "Name", "") or ""),
                        "owner": str(getattr(r, "Owner", "") or ""),
                        "mine": int(getattr(r, "IsMine", 0) or 0) == 1})
    except Exception:
        pass
    return out


def _scope_ids(query_name):
    """Resolve a saved search to people ids. Returns (ids, note, error)."""
    if not query_name:
        return (None, "", "")
    try:
        ids = list(q.QueryPeopleIds(query_name))
    except Exception as e:
        return (None, "", 'Could not run the search "%s": %s'
                % (query_name, str(e)))
    if not ids:
        return ([], "", 'The search "%s" matched nobody, or no longer exists.'
                % query_name)
    note = ""
    if len(ids) > CUSTOM_MAX_IDS:
        # Said out loud. A tile that quietly charts the first 5,000 of 9,000
        # people is worse than one that admits it.
        note = ("Limited to the first %d of %d people in this search."
                % (CUSTOM_MAX_IDS, len(ids)))
        ids = ids[:CUSTOM_MAX_IDS]
    return (ids, note, "")


def sql_literal(v):
    """A string as a SQL literal. Values here come from our own query results,
    but they still get quoted and escaped: the day someone passes a label in
    from somewhere else, this is what stops it being an injection."""
    return "N'" + unicode(v).replace("'", "''") + "'"


def drill_custom_tile(spec, label, global_range=None, global_query=None):
    dom_cat = DOMAIN_CATEGORY.get(str(spec.get("domain", "")), "")
    if dom_cat and not can_view_category(dom_cat):
        return {"success": False, "restricted": True,
                "error": "You do not have access to this data."}
    """The people behind one bar, slice or row of a custom tile.

    Same base, same filters, same scope as the tile itself, restricted to the
    one group. Anything else and the drill would disagree with the chart it
    came from.
    """
    dom = CUSTOM_DOMAINS.get(str(spec.get("domain", "")))
    if not dom:
        return {"success": False, "error": "Unknown data type."}
    dim_key = str(spec.get("dimension", "none")) or "none"
    dim = dom["dimensions"].get(dim_key)
    if not dim or dim["expr"] is None:
        return {"success": False,
                "error": "This tile is a single total, so there is nothing to open."}

    ids, note, err = _scope_ids(str(spec.get("query", "") or "")
                                or str(global_query or ""))
    if err:
        return {"success": False, "error": err}

    where = [dom["where"]] + custom_filter_sql(spec)
    if ids is not None:
        if not ids:
            return {"success": True, "people": [], "note": note}
        where.append("%s IN (%s)" % (dom["person_col"],
                                     ",".join(str(int(i)) for i in ids)))
    if dom["dates"]:
        lo, hi = date_range_sql(global_range) if global_range else (None, None)
        if lo:
            where.append("%s >= %s" % (dom["date_col"], lo))
            if hi:
                where.append("%s < %s" % (dom["date_col"], hi))
        else:
            try:
                months = int(spec.get("months", 12))
            except Exception:
                months = 12
            months = max(1, min(months, 120))
            where.append("%s >= DATEADD(MONTH, -%d, GETDATE())"
                         % (dom["date_col"], months))
            where.append("%s < DATEADD(DAY, 1, GETDATE())" % dom["date_col"])
    where.append("%s = %s" % (dim["expr"], sql_literal(label)))

    sql = ("SELECT DISTINCT TOP 500 p.PeopleId, p.Name2 AS Nm, p.Age AS Ag,"
           " ISNULL(p.EmailAddress,'') AS Em %s %s WHERE %s ORDER BY p.Name2"
           % (dom["from"], dim.get("join", "") or "", " AND ".join(where)))
    try:
        rows = q.QuerySql(sql)
    except Exception as e:
        return {"success": False, "error": "Drill failed: " + str(e)}

    people = []
    for r in rows:
        people.append({"id": int(getattr(r, "PeopleId", 0) or 0),
                       "name": str(getattr(r, "Nm", "") or ""),
                       "age": (getattr(r, "Ag", None)
                               if getattr(r, "Ag", None) is not None else ""),
                       "email": str(getattr(r, "Em", "") or "")})
    capped = len(people) >= 500
    return {"success": True, "people": people, "note": note,
            "capped": capped, "label": label,
            "dimension_label": dim["label"]}


def run_custom_tile(spec, global_range=None, global_query=None):
    """Execute a custom tile spec. Every SQL fragment comes from the catalog."""
    dom_key = str(spec.get("domain", ""))
    dom = CUSTOM_DOMAINS.get(dom_key)
    if not dom:
        return {"success": False, "error": "Unknown data type: " + dom_key}
    dom_cat = DOMAIN_CATEGORY.get(dom_key, "")
    if dom_cat and not can_view_category(dom_cat):
        return {"success": False, "restricted": True,
                "error": "You do not have access to this data."}

    meas = dom["measures"].get(str(spec.get("measure", "")))
    if not meas:
        return {"success": False, "error": "Unknown measure for " + dom["label"]}

    dim_key = str(spec.get("dimension", "none")) or "none"
    dim = dom["dimensions"].get(dim_key)
    if not dim:
        return {"success": False, "error": "Unknown grouping for " + dom["label"]}

    ids, note, err = _scope_ids(str(spec.get("query", "") or "")
                                or str(global_query or ""))
    if err:
        return {"success": False, "error": err}

    where = [dom["where"]] + custom_filter_sql(spec)
    if ids is not None:
        if not ids:
            return {"success": True, "rows": [], "note": note,
                    "money": meas["money"], "empty": True}
        where.append("%s IN (%s)" % (dom["person_col"],
                                     ",".join(str(int(i)) for i in ids)))

    if dom["dates"]:
        lo, hi = date_range_sql(global_range) if global_range else (None, None)
        if lo:
            where.append("%s >= %s" % (dom["date_col"], lo))
            if hi:
                where.append("%s < %s" % (dom["date_col"], hi))
        else:
            try:
                months = int(spec.get("months", 12))
            except Exception:
                months = 12
            months = max(1, min(months, 120))
            where.append("%s >= DATEADD(MONTH, -%d, GETDATE())"
                         % (dom["date_col"], months))
            where.append("%s < DATEADD(DAY, 1, GETDATE())" % dom["date_col"])

    join = dim.get("join", "") or ""
    if dim["expr"] is None:
        sql = "SELECT %s AS Value %s %s WHERE %s" % (
            meas["expr"], dom["from"], join, " AND ".join(where))
    else:
        order = dim.get("order")
        if dim_key == "month":
            order_by = "ORDER BY 1"
        elif order:
            order_by = "ORDER BY MIN(%s)" % order
        else:
            order_by = "ORDER BY 2 DESC"
        sql = ("SELECT TOP 60 %s AS Label, %s AS Value %s %s WHERE %s"
               " GROUP BY %s %s"
               % (dim["expr"], meas["expr"], dom["from"], join,
                  " AND ".join(where), dim["expr"], order_by))

    try:
        rows = q.QuerySql(sql)
    except Exception as e:
        return {"success": False, "error": "Query failed: " + str(e)}

    out = []
    for r in rows:
        val = getattr(r, "Value", None)
        try:
            val = float(val) if val is not None else 0.0
        except:
            # Bare on purpose: a .NET Decimal or DateTime here raises something
            # that "except Exception" can miss in IronPython.
            val = 0.0
        if dim["expr"] is None:
            out.append({"label": meas["label"], "value": val})
        else:
            out.append({"label": str(getattr(r, "Label", "") or "(none)"),
                        "value": val})
    return {"success": True, "rows": out, "note": note,
            "money": meas["money"], "single": dim["expr"] is None,
            "measure_label": meas["label"],
            "dimension_label": dim["label"]}


# ---------------------------------------------------------------------------
# AJAX
# ---------------------------------------------------------------------------
def handle_ajax(action, uid):
    if action == "list_dashboards":
        return safe_json({"success": True,
                          "dashboards": [{"id": d.get("id"), "name": d.get("name", ""),
                                          "icon": d.get("icon", "fa-th-large"),
                                          "description": d.get("description", ""),
                                          "scope": d.get("scope", "personal"),
                                          "tabs": len(d.get("tabs") or [])}
                                         for d in visible_dashboards(uid)],
                          "can_share": can_edit_shared()})

    if action == "list_library":
        status = library_status(uid)
        lib = []
        for l in get_library():
            ok, why = library_available(l)
            mine = (status.get(l["id"], {}) or {}).get("copies", [])
            lib.append({"id": l["id"], "name": l["name"],
                        # Installed copies, so a card can say so and offer to
                        # open one instead of silently making a duplicate.
                        "installed": mine,
                        "icon": l.get("icon", ""),
                        "description": l.get("description", ""),
                        "audience": l.get("audience", ""),
                        "needs": library_needs(l),
                        "usable": ok, "why": why,
                        "tabs": len(l.get("tabs") or []),
                        "tiles": sum(len(t.get("tiles") or [])
                                     for t in l.get("tabs") or [])})
        return safe_json({"success": True, "library": lib,
                          "reports_script": reports_script_name()})

    if action == "get_settings":
        return safe_json({"success": True,
                          "reports_script": reports_script_name(),
                          "configured": str((load_settings() or {})
                                            .get("reports_script", "") or ""),
                          "fiscal_month": int((load_settings() or {})
                                              .get("fiscal_year_start_month", 10)
                                              or 10),
                          "bg_check_days": int((load_settings() or {})
                                               .get("bg_check_days", 730) or 730),
                          "attendance_scope": str((load_settings() or {})
                                                  .get("attendance_scope", "waag")
                                                  or "waag"),
                          "bg_approvals": BG_APPROVALS,
                          "bg_approval_ok": bg_approval_list()[0],
                          "bg_approval_set": bg_approval_list()[1],
                          "can_share": can_edit_shared()})

    if action == "save_settings":
        if not can_edit_shared():
            return safe_json({"success": False,
                              "error": "You do not have permission to change this."})
        st = load_settings()
        st["reports_script"] = get_param("reports_script", "").strip()
        fmv = get_param("fiscal_month", "").strip()
        if fmv.isdigit() and 1 <= int(fmv) <= 12:
            st["fiscal_year_start_month"] = int(fmv)
        bgv = get_param("bg_check_days", "").strip()
        if bgv.isdigit() and 30 <= int(bgv) <= 3650:
            st["bg_check_days"] = int(bgv)
        asc = get_param("attendance_scope", "").strip()
        if asc in ("waag", "all"):
            st["attendance_scope"] = asc
        appr = [x.strip() for x in get_param("bg_approval_ok", "").split(",")
                if x.strip() in BG_APPROVALS]
        if appr:
            st["bg_approval_ok"] = ",".join(appr)
        save_settings(st)
        _REPORTS_SCRIPT_CACHE[0] = None
        return safe_json({"success": True, "reports_script": reports_script_name()})

    if action == "upload_begin":
        if not can_edit_shared():
            return safe_json({"success": False, "error": "Permission required."})
        save_content_json(UPLOAD_KEY, {"name": get_param("s_name", ""),
                                       "chunks": [], "expect": 0})
        return safe_json({"success": True})

    if action == "upload_chunk":
        if not can_edit_shared():
            return safe_json({"success": False, "error": "Permission required."})
        st = load_content_json(UPLOAD_KEY, None)
        if not isinstance(st, dict):
            return safe_json({"success": False,
                              "error": "Upload was not started."})
        st.setdefault("chunks", []).append(get_param("s_data", ""))
        save_content_json(UPLOAD_KEY, st)
        return safe_json({"success": True, "received": len(st["chunks"])})

    if action == "upload_finish":
        if not can_edit_shared():
            return safe_json({"success": False, "error": "Permission required."})
        st = load_content_json(UPLOAD_KEY, None)
        if not isinstance(st, dict) or not st.get("chunks"):
            return safe_json({"success": False, "error": "Nothing was uploaded."})
        name = (get_param("s_name", "") or st.get("name", "")).strip()
        if not name:
            return safe_json({"success": False, "error": "No script name given."})
        # Entity-decoding happens ONCE, over the whole reassembled body. Doing it
        # per chunk could split an entity across a boundary and corrupt it.
        body = decode_payload("".join(st.get("chunks", [])))
        if len(body) < 1000:
            return safe_json({"success": False,
                              "error": "Body is only %d characters; refusing to "
                                       "overwrite a script with that."
                                       % len(body)})
        try:
            model.WriteContentPython(name, body)
        except Exception as e:
            return safe_json({"success": False, "error": "Write failed: " + str(e)})
        save_content_json(UPLOAD_KEY, {})
        return safe_json({"success": True, "name": name, "chars": len(body)})

    # Tagging and tasking live here rather than being forwarded to Enterprise
    # Reporting: a catalog tile is meant to work without that script installed,
    # and an action bar that only works on some tiles is worse than none.
    # Answers to a keyword's extra questions. TouchPoint writes these only
    # through TaskNoteEdit or TaskNoteComplete, and completing a brand new task
    # is wrong, so the edit path is used. Two hazards are handled below:
    # TaskNoteEdit deletes every keyword before re-reading them from
    # IncomingKeywords, and it only touches the answers when that list contains
    # a keyword that actually has questions.
    def write_extra_answers(task_id, kw_ids, answers):
        """Returns (written, error). Never raises: a task already exists by the
        time this runs, and losing it to a failed answer write would be worse
        than a task with unanswered questions."""
        if not answers or not kw_ids:
            return 0, None
        try:
            import clr
            clr.AddReference("CmsData")
            from CmsData import TNDropdownList
        except Exception as e:
            return 0, "cannot reach CmsData types: " + str(e)
        try:
            qs = list(model.GetExtraQuestionsByKeywords(kw_ids))
            filled = 0
            for qq in qs:
                key = str(qq.keywordExtraValueId)
                if key not in answers:
                    continue
                val = answers[key]
                dt = int(qq.dataType)
                if dt in (3, 4, 8):          # text, multiline, date
                    qq.response = str(val)
                    filled += 1
                elif dt == 6:                # yes/no
                    qq.boolResponse = str(val).lower() in ("true", "1", "yes")
                    filled += 1
                elif dt == 5:                # dropdown: hand back an option
                    for opt in (qq.dropdownOptions or []):
                        if str(opt.Value) == str(val):
                            qq.dropdownResponse = opt
                            filled += 1
                            break
                elif dt == 7:                # checkboxes
                    picked = set(str(x) for x in (val or []))
                    for cb in (qq.checkboxesList or []):
                        cb.boolResponse = str(cb.keywordExtraValueOptionId) in picked
                    filled += 1
            if not filled:
                return 0, None

            incoming = []
            for k in kw_ids:
                d = TNDropdownList()
                d.Value = int(k)
                d.Text = str(k)
                incoming.append(d)

            vm = model.GetTaskNote(task_id)
            vm.IncomingKeywords = incoming
            vm.extraQuestions = qs
            model.TaskNoteEdit(vm)
        except Exception as e:
            # Put the keywords back regardless: the edit may have deleted them
            # before failing.
            try:
                model.SetTaskNoteKeywordIds(task_id, kw_ids)
            except:
                pass
            return 0, str(e)

        # The edit rebuilt the keyword rows from our list; re-assert them so a
        # partial run cannot leave the task without its keywords.
        try:
            model.SetTaskNoteKeywordIds(task_id, kw_ids)
        except:
            pass
        # Counted from the database rather than assumed, so a silent no-op
        # shows up as zero instead of a success message.
        try:
            chk = q.QuerySql("""
                SELECT COUNT(*) AS N FROM TaskNoteExtraValue WITH (NOLOCK)
                WHERE TaskNoteId = %d AND Response IS NOT NULL
                  AND LEN(LTRIM(RTRIM(Response))) > 0
            """ % int(task_id))
            return int(list(chk)[0].N or 0), None
        except Exception as e:
            return 0, str(e)

    if action == "filter_options":
        rep = catalog_report(get_param("report_id", ""))
        if not rep:
            return safe_json({"success": False, "error": "Report not installed."})
        want = [x.strip() for x in get_param("names", "").split(",") if x.strip()]
        out = {}
        for prm in (rep.get("parameters") or []):
            if prm.get("name") not in want:
                continue
            src = prm.get("source_sql", "") or ""
            if not src:
                continue
            # The option list is a query the definition carries. It must be as
            # read-only as the report itself.
            ok, why = sql_is_read_only(src)
            if not ok:
                continue
            try:
                rows = q.QuerySql(src)
            except Exception:
                continue
            vals = []
            for r in rows:
                v = getattr(r, "value", None)
                l = getattr(r, "label", None)
                if v is None:
                    continue
                vals.append({"value": str(v), "label": str(l if l is not None else v)})
                if len(vals) >= 300:
                    break
            out[prm.get("name")] = vals
        return safe_json({"success": True, "options": out})

    if action == "keyword_questions":
        kw = [int(x) for x in get_param("keyword_ids", "").split(",")
              if x.strip().isdigit()]
        if not kw:
            return safe_json({"success": True, "questions": []})
        rows = q.QuerySql("""
            SELECT kev.KeywordExtraValueId, kev.KeywordId, kev.DataType,
                   kev.SortOrder, ISNULL(kev.Name, '') AS Name,
                   ISNULL(k.Description, k.Code) AS Keyword
            FROM KeywordExtraValue kev WITH (NOLOCK)
            JOIN Keyword k WITH (NOLOCK) ON k.KeywordId = kev.KeywordId
            WHERE kev.KeywordId IN (%s)
            ORDER BY k.Description, kev.SortOrder, kev.KeywordExtraValueId
        """ % ",".join(str(x) for x in kw))
        qs = []
        for r in rows:
            qs.append({"id": int(r.KeywordExtraValueId),
                       "keyword": r.Keyword or "",
                       "type": int(r.DataType or 0),
                       "name": r.Name or "",
                       "options": []})
        if qs:
            opts = q.QuerySql("""
                SELECT o.KeywordExtraValueOptionId, o.KeywordExtraValueId,
                       ISNULL(o.Name, '') AS Name
                FROM KeywordExtraValueOption o WITH (NOLOCK)
                WHERE o.KeywordExtraValueId IN (%s)
                ORDER BY o.Name
            """ % ",".join(str(x["id"]) for x in qs))
            by_id = {}
            for x in qs:
                by_id[x["id"]] = x
            for o in opts:
                tgt = by_id.get(int(o.KeywordExtraValueId))
                if tgt is not None:
                    tgt["options"].append({"id": int(o.KeywordExtraValueOptionId),
                                           "name": o.Name})
        return safe_json({"success": True, "questions": qs})

    if action == "task_keywords":
        # Header and instruction rows are display furniture in TouchPoint's own
        # task form, not questions, so they are not counted as such here.
        rows = q.QuerySql("""
            SELECT k.KeywordId, k.Code, k.Description,
                   (SELECT COUNT(*) FROM KeywordExtraValue kev WITH (NOLOCK)
                     WHERE kev.KeywordId = k.KeywordId
                       AND kev.DataType NOT IN (1, 2)) AS Questions,
                   (SELECT TOP 1 STUFF((
                        SELECT ', ' + CAST(k2.Name AS VARCHAR(MAX))
                        FROM KeywordExtraValue k2 WITH (NOLOCK)
                        WHERE k2.KeywordId = k.KeywordId
                          AND k2.DataType NOT IN (1, 2)
                        ORDER BY k2.SortOrder
                        FOR XML PATH('')), 1, 2, '')) AS QuestionNames
            FROM Keyword k WITH (NOLOCK)
            WHERE k.IsActive = 1
            ORDER BY k.Description
        """)
        out = []
        for r in rows:
            out.append({"id": r.KeywordId,
                        "code": r.Code or "",
                        "name": r.Description or r.Code or "",
                        "questions": int(r.Questions or 0),
                        "question_names": r.QuestionNames or ""})
        return safe_json({"success": True, "keywords": out})

    if action == "search_assignee":
        term = get_param("search_term", "").strip()
        if len(term) < 2:
            return safe_json({"success": True, "people": []})
        like = term.replace("'", "''").replace("%", "").replace("_", "")
        # Matched against every way a person might be typed, not just Name2
        # ("Last, First"), so "Ben Swaby" works as well as "Swaby, Ben".
        # Username and email come back too: when one person has several
        # accounts, age cannot tell them apart but the login can.
        pid_term = like if like.isdigit() else "-1"
        # Shape matters more than the conditions here. Putting the username
        # test in a correlated EXISTS inside the OR made SQL Server walk 11k
        # Users rows for each of 55k People: 46 seconds. Collapsing Users to
        # one row per person FIRST, joining to it, and building the username
        # list only for the 25 rows that survive runs the same search in 79ms.
        rows = q.QuerySql("""
            WITH lg AS (
                SELECT u.PeopleId,
                       MAX(CASE WHEN u.Username LIKE '%%%s%%' THEN 1 ELSE 0 END)
                           AS UserHit
                FROM dbo.Users u WITH (NOLOCK)
                GROUP BY u.PeopleId
            ), hits AS (
                SELECT TOP 25 p.PeopleId, p.Name2,
                       ISNULL(p.Age, -1) AS Age,
                       ISNULL(p.EmailAddress, '') AS Em
                FROM dbo.People p WITH (NOLOCK)
                JOIN lg ON lg.PeopleId = p.PeopleId
                WHERE p.IsDeceased = 0 AND p.ArchivedFlag = 0
                  AND (lg.UserHit = 1
                       OR p.PeopleId = %s
                       OR p.Name2 LIKE '%%%s%%'
                       OR (ISNULL(p.PreferredName, p.FirstName) + ' '
                           + p.LastName) LIKE '%%%s%%'
                       OR (p.FirstName + ' ' + p.LastName) LIKE '%%%s%%'
                       OR ISNULL(p.EmailAddress, '') LIKE '%%%s%%')
                ORDER BY p.Name2
            )
            SELECT h.PeopleId, h.Name2, h.Age, h.Em,
                   ISNULL(STUFF((SELECT ', ' + u2.Username
                                 FROM dbo.Users u2 WITH (NOLOCK)
                                 WHERE u2.PeopleId = h.PeopleId
                                 FOR XML PATH('')), 1, 2, ''), '') AS Unames
            FROM hits h
            ORDER BY h.Name2
        """ % (like, pid_term, like, like, like, like))
        return safe_json({"success": True,
                          "people": [{"id": r.PeopleId, "name": r.Name2,
                                      "age": int(r.Age),
                                      "email": str(r.Em or ""),
                                      "user": str(r.Unames or "")}
                                     for r in rows]})

    if action in ("bulk_tag", "bulk_task"):
        pids = []
        for x in get_param("people_ids", "").split(","):
            x = x.strip()
            if x.isdigit():
                pids.append(int(x))
        # Bounded because these run one model call per person.
        if not pids:
            return safe_json({"success": False, "error": "Nobody was selected."})
        if len(pids) > 500:
            return safe_json({"success": False,
                              "error": "That is %d people. Select 500 or fewer."
                                       % len(pids)})
        me = current_user_id()

        if action == "bulk_tag":
            name = get_param("tag_name", "").strip()
            if not name:
                return safe_json({"success": False, "error": "Name the tag."})
            try:
                model.AddTag("peopleids='%s'" % ",".join(str(x) for x in pids),
                             name, me, False)
            except Exception as e:
                return safe_json({"success": False, "error": str(e)})
            return safe_json({"success": True, "count": len(pids),
                              "message": "Tagged %d people with '%s'."
                                         % (len(pids), name)})

        msg = get_param("task_message", "").strip()
        if not msg:
            return safe_json({"success": False, "error": "Say what the task is."})
        assignee = get_param("assignee_id", "").strip()
        assignee = int(assignee) if assignee.isdigit() else me
        kws = [int(x) for x in get_param("keyword_ids", "").split(",")
               if x.strip().isdigit()]
        answers = {}
        raw_a = decode_payload(get_param("answers", ""))
        if raw_a:
            try:
                answers = json.loads(raw_a)
            except Exception:
                answers = {}
        due = get_param("due_date", "").strip()
        try:
            due = model.ParseDate(due) if due else None
        except:
            due = None
        made, failed, answered = 0, 0, 0
        ans_err = None
        for pid in pids:
            try:
                # 9th argument is the keyword list. Passing [] here, which is
                # what the reporting script does, is why tasks were arriving
                # with no keyword at all.
                tid = model.CreateTaskNote(me, pid, assignee, None, False, msg,
                                           "", due, kws, True)
                made += 1
                if answers and tid:
                    got, err = write_extra_answers(tid, kws, answers)
                    if got:
                        answered += 1
                    elif err and not ans_err:
                        ans_err = err
            except:
                failed += 1
        # Reported rather than swallowed: a partial run looks identical to a
        # clean one otherwise.
        note = ""
        if answers:
            if answered == made and made:
                note = " Answers saved."
            elif answered:
                note = " Answers saved on %d of them." % answered
            else:
                # Said out loud: the task exists but the questions are blank,
                # and finding that out later on the task itself is worse.
                note = (" Answers could NOT be saved"
                        + ((": " + ans_err) if ans_err else "")
                        + " - open the task to fill them in.")
        return safe_json({"success": True, "count": made, "failed": failed,
                          "answered": answered,
                          "message": "Created %d task%s.%s%s"
                                     % (made, "" if made == 1 else "s",
                                        (" %d could not be created."
                                         % failed) if failed else "", note)})

    if action == "ops_list":
        items, err = fetch_ops_catalog()
        if err:
            return safe_json({"success": False, "error": err})
        out = []
        for it in items:
            out.append({"id": it["id"], "name": it["name"], "cat": it["cat"],
                        "group": it["group"], "freq": it["freq"],
                        "note": it["note"], "steps": it["steps"]})
        return safe_json({"success": True, "checks": out})

    if action == "ops_run":
        return safe_json(run_ops_check(get_param("check_id", "")))

    if action == "catalog_browse":
        items, err = fetch_published_catalog()
        if err:
            return safe_json({"success": False, "error": err})
        dash = []
        reports = []
        if isinstance(items, dict):
            reports = items.get("reports", []) or []
            dash = items.get("dashboards", []) or []
        else:
            reports = items or []
        have = {}
        for r in load_report_catalog().get("reports", []):
            have[str(r.get("id", ""))] = r
        out_r = []
        for it in reports:
            # Same rule as execution, so the catalog does not advertise
            # reports this user would only be refused.
            if not can_view_category(it.get("category", "")):
                continue
            rid = str(it.get("id", ""))
            cur = have.get(rid)
            out_r.append({
                "id": rid, "name": it.get("name", ""),
                "description": it.get("description", ""),
                "category": it.get("category", "custom"),
                "tokens": it.get("tokens", []) or [],
                "lookups": it.get("lookups", []) or [],
                # Derived from the SQL rather than the upstream flag. The
                # flag is both over- and under-inclusive: one entry declares
                # itself scopeable while aliasing People as something else,
                # and fourteen that take a scope perfectly well never declare
                # it. What the SQL can actually accept is the honest answer.
                "scopeable": accepts_people_scope(it),
                # Read from the INSTALLED copy when there is one. A setting
                # only reaches a report whose installed SQL actually carries
                # the token, so promising otherwise would be a lie: a stale
                # copy with 730 baked in ignores the setting entirely.
                "settings": setting_notes(
                    sql_tokens(cur.get("sql_template", "")) if cur
                    else (it.get("tokens") or [])),
                # Column names, so the tile can be configured without first
                # running the report. Parsed from the SELECT, with the
                # placeholders stripped so an unresolved token is not mistaken
                # for a column.
                "columns": sql_select_columns(resolve_year_tokens(
                    strip_injection_tokens(it.get("sql_template", "") or ""))),
                "display": resolve_display(it.get("display", {}) or {}),
                # The filters the report declares, so a tile can be narrowed
                # instead of always showing the whole report.
                "params": [{"name": x.get("name", ""),
                            "label": x.get("label", x.get("name", "")),
                            "type": x.get("type", "text"),
                            "default": x.get("default", "")}
                           for x in (it.get("parameters") or [])
                           if x.get("name")],
                "installed": cur is not None,
                # A hash change is the only reliable signal that upstream
                # edited a report we already hold.
                "outdated": bool(cur and cur.get("hash") != it.get("hash", "")),
            })
        out_d = []
        for d in dash:
            out_d.append({
                "id": d.get("id", ""), "name": d.get("name", ""),
                "description": d.get("description", ""),
                "audience": d.get("audience", ""),
                "tab_count": d.get("tab_count", len(d.get("tabs") or [])),
                "tile_count": d.get("tile_count", 0),
                "requires_reports": d.get("requires_reports", []) or [],
                "requires_widgets": d.get("requires_widgets", []) or [],
            })
        return safe_json({"success": True, "reports": out_r, "dashboards": out_d,
                          "installed": len(have)})

    if action == "catalog_install":
        if not can_edit_shared():
            return safe_json({"success": False,
                              "error": "You do not have permission to install "
                                       "catalog reports."})
        items, err = fetch_published_catalog()
        if err:
            return safe_json({"success": False, "error": err})
        reports = items.get("reports", []) if isinstance(items, dict) else (items or [])
        by_id = dict((str(r.get("id", "")), r) for r in reports)

        wanted = [x.strip() for x in get_param("ids", "").split(",") if x.strip()]
        # Installing a dashboard pulls the reports it needs. A board of dead
        # tiles is worse than refusing the install.
        dash_id = get_param("dash_id", "")
        dash_obj = None
        if dash_id:
            for d in (items.get("dashboards", []) if isinstance(items, dict) else []):
                if str(d.get("id", "")) == dash_id:
                    dash_obj = d
                    for rid in (d.get("requires_reports", []) or []):
                        if rid not in wanted:
                            wanted.append(rid)

        store = load_report_catalog()
        have = dict((str(r.get("id", "")), i)
                    for i, r in enumerate(store.get("reports", [])))
        installed, missing = [], []
        for rid in wanted:
            src = by_id.get(rid)
            if not src:
                missing.append(rid)
                continue
            # Checked again here, not only at publish: what was safe when it was
            # published is not necessarily what arrived.
            ok, why = sql_is_read_only(re.sub(r"\{[a-z_0-9]+\}", "",
                                             src.get("sql_template", "") or ""))
            if not ok:
                missing.append(rid + " (unsafe SQL: " + why + ")")
                continue
            entry = dict(src)
            entry["installed_at"] = str(model.DateTime) if hasattr(model, "DateTime") else ""
            if rid in have:
                store["reports"][have[rid]] = entry
            else:
                store.setdefault("reports", []).append(entry)
            installed.append(rid)
        save_content_json(CONTENT_REPORTS, store)

        made = ""
        if dash_obj:
            tabs, miss_w = resolve_widget_names(
                json.loads(safe_json(dash_obj.get("tabs", []))))
            for tab in tabs:
                for t in (tab.get("tiles") or []):
                    if t.get("kind") == "report":
                        t["source"] = "catalog"
            d = {"id": "", "name": dash_obj.get("name", "Dashboard"),
                 "description": dash_obj.get("description", ""),
                 "icon": "fa-th-large", "roles": [], "scope": "personal",
                 "from_catalog": dash_id, "tabs": tabs}
            ok2, res2 = store_dashboard(uid, d)
            if ok2:
                made = res2 if not isinstance(res2, dict) else res2.get("id", "")
        return safe_json({"success": True, "installed": installed,
                          "missing": missing, "dashboard_id": made})

    if action == "run_catalog":
        ids = None
        qname = get_param("q_name", "")
        if qname:
            got, note, err = _scope_ids(qname)
            if err:
                return safe_json({"success": False, "error": err})
            ids = got
        vals = {}
        raw = decode_payload(get_param("filters", ""))
        if raw:
            try:
                vals = json.loads(raw)
            except Exception:
                vals = {}
        serv = [x.strip() for x in get_param("serving_types", "").split(",")
                if x.strip().isdigit()]
        return safe_json(run_catalog_report(get_param("report_id", ""), vals,
                                            ids, serv,
                                            get_param("global_range", ""),
                                            get_param("global_query", ""),
                                            get_param("global_program", "")))

    if action == "admin_dashboards":
        if not can_manage_users():
            return safe_json({"success": False,
                              "error": "This needs the Admin or Developer role."})

        def _counts(d):
            t = 0
            for tab in (d.get("tabs") or []):
                t += len(tab.get("tiles") or [])
            return len(d.get("tabs") or []), t

        rows = []
        for o in all_dashboard_owners():
            for d in load_personal(o["uid"]).get("dashboards", []):
                tabs_n, tiles_n = _counts(d)
                rows.append({"uid": o["uid"], "owner": o["owner"],
                             "scope": "personal", "id": d.get("id", ""),
                             "name": d.get("name", ""), "tabs": tabs_n,
                             "tiles": tiles_n, "mine": (o["uid"] == uid)})
        # Shared dashboards live in one store rather than a per-person slot, so
        # they are not reachable through the owner loop above. Listing them here
        # means this screen is every dashboard, not just the personal ones.
        shared = load_shared().get("dashboards", [])
        want = [d.get("owner") for d in shared if str(d.get("owner", "")).strip()]
        names = people_names(want)
        for d in shared:
            tabs_n, tiles_n = _counts(d)
            ow = d.get("owner")
            rows.append({"uid": 0, "scope": "shared", "id": d.get("id", ""),
                         "name": d.get("name", ""), "tabs": tabs_n,
                         "tiles": tiles_n, "mine": (ow == uid),
                         "roles": ", ".join(d.get("roles") or []),
                         "owner": names.get(ow, "") or "Shared"})
        rows.sort(key=lambda x: ((x.get("owner") or "").lower(),
                                 (x.get("name") or "").lower()))
        return safe_json({"success": True, "rows": rows,
                          "can_delete_shared": can_edit_shared(),
                          "legacy": legacy_rows()})

    if action == "legacy_assign":
        idxs = []
        for x in get_param("idxs", get_param("idx", "")).split(","):
            x = x.strip()
            if x.lstrip("-").isdigit():
                idxs.append(int(x))
        try:
            to_uid = int(get_param("to_uid", "0") or 0)
        except Exception:
            to_uid = 0
        ok, msg = adopt_many(idxs, to_uid)
        return safe_json({"success": ok, "message": msg if ok else "",
                          "error": "" if ok else msg})

    if action == "reassign":
        refs = [x.strip() for x in get_param("refs", "").split(",") if x.strip()]
        try:
            to_uid = int(get_param("to_uid", "0") or 0)
        except Exception:
            to_uid = 0
        ok, msg = reassign_dashboards(refs, to_uid)
        return safe_json({"success": ok, "message": msg if ok else "",
                          "error": "" if ok else msg})

    if action == "legacy_drop":
        try:
            idx = int(get_param("idx", "-1"))
        except Exception:
            idx = -1
        ok, msg = drop_one(idx)
        return safe_json({"success": ok, "message": msg if ok else "",
                          "error": "" if ok else msg})

    if action == "apply_update":
        if not can_apply_update():
            return safe_json({"success": False,
                              "error": "Updating this script needs the Admin "
                                       "or Developer role."})
        url = DC_API_WORKER + "/scripts/" + DC_SCRIPT_ID
        try:
            code = model.RestGet(url, dc_headers())
        except Exception as fe:
            return safe_json({"success": False,
                              "error": "Could not reach the update server: "
                                       + str(fe) + " Install by hand from "
                                       + url + " if this keeps happening."})
        code = str(code or "")
        # A Cloudflare challenge or an error page is still a 200 with a body.
        # Writing one of those over a working script has no undo.
        if len(code) < 2000 or "DC_SCRIPT_ID" not in code \
                or "TPxi_Dashboards" not in code:
            return safe_json({"success": False,
                              "error": "What came back does not look like this "
                                       "script (" + str(len(code)) + " bytes), "
                                       "so nothing was changed. Check " + url})
        # The version the NEW file declares. Used below to prove the write
        # landed: checking for the version we are already running would pass
        # even when nothing was replaced.
        m = re.search(r'APP_VERSION\s*=\s*["\']([0-9][0-9.]*)["\']', code)
        new_ver = m.group(1) if m else ""
        target = get_script_name() or DC_SCRIPT_ID
        try:
            model.WriteContentPython(target, code)
        except Exception as we:
            return safe_json({"success": False,
                              "error": "Could not write '" + target + "': "
                                       + str(we)})
        # WriteContentPython is an upsert with no role check, and it happily
        # CREATES a slot. If this script is installed under a name we did not
        # detect, the write succeeded but landed beside the running copy.
        try:
            back = str(model.PythonContent(target) or "")
        except Exception:
            back = ""
        landed = False
        if new_ver:
            landed = bool(re.search(
                r'APP_VERSION\s*=\s*["\']' + re.escape(new_ver) + r'["\']',
                back))
        if new_ver and not landed:
            return safe_json({"success": False,
                              "error": "Wrote to '" + target + "' but v"
                                       + new_ver + " is not there afterwards. "
                                       "This script is probably installed under "
                                       "a different name: find it under Admin > "
                                       "Advanced > Special Content > Python and "
                                       "paste the code from " + url})
        return safe_json({"success": True, "name": target, "version": new_ver})

    if action == "diagnose_reports":
        # TouchPoint resolves a script by NAME with FirstOrDefault, no ordering
        # and no archived filter, so a second row with the same name can be the
        # one that actually executes while you edit the other. That is invisible
        # from the editor, so look at the rows directly.
        want = reports_script_name() or "EnterpriseReporting"
        safe = want.replace("'", "''")
        rows = []
        try:
            for r in q.QuerySql("""
                SELECT Id, Name, TypeID,
                       -- Archived is a DATETIME, not a bit. ISNULL(...,0) hands
                       -- back a date, and int() on a .NET DateTime throws.
                       CASE WHEN Archived IS NULL THEN 0 ELSE 1 END AS Arch,
                       LEN(CAST(Body AS VARCHAR(MAX))) AS Bytes,
                       CASE WHEN CAST(Body AS VARCHAR(MAX)) LIKE '%%REPORTING_BUILD%%'
                            THEN 1 ELSE 0 END AS HasBuild,
                       CASE WHEN CAST(Body AS VARCHAR(MAX)) LIKE '%%list_reports%%'
                            THEN 1 ELSE 0 END AS HasAction,
                       CASE WHEN CAST(Body AS VARCHAR(MAX)) LIKE '%%load_person_detail%%'
                            THEN 1 ELSE 0 END AS M1,
                       CASE WHEN CAST(Body AS VARCHAR(MAX)) LIKE '%%save_contact_methods%%'
                            THEN 1 ELSE 0 END AS M2,
                       CASE WHEN CAST(Body AS VARCHAR(MAX)) LIKE '%%eng_person_engagement_scorecard%%'
                            THEN 1 ELSE 0 END AS M3,
                       CASE WHEN CAST(Body AS VARCHAR(MAX)) LIKE '%%serving_types%%'
                            THEN 1 ELSE 0 END AS M4
                FROM dbo.Content WITH (NOLOCK)
                WHERE Name = '%s' OR Name LIKE '%%%s%%'
                ORDER BY TypeID, Id
            """ % (safe, safe)):
                rows.append({"id": int(r.Id),
                             "name": str(getattr(r, "Name", "") or ""),
                             "type": int(getattr(r, "TypeID", 0) or 0),
                             "archived": int(getattr(r, "Arch", 0) or 0),
                             "bytes": int(getattr(r, "Bytes", 0) or 0),
                             "has_build": int(getattr(r, "HasBuild", 0) or 0),
                             "has_action": int(getattr(r, "HasAction", 0) or 0),
                             "m1": int(getattr(r, "M1", 0) or 0),
                             "m2": int(getattr(r, "M2", 0) or 0),
                             "m3": int(getattr(r, "M3", 0) or 0),
                             "m4": int(getattr(r, "M4", 0) or 0)})
        except Exception as e:
            return safe_json({"success": False, "error": str(e)})
        return safe_json({"success": True, "looking_for": want, "rows": rows})

    if action == "member_types":
        out = []
        try:
            for r in q.QuerySql("""
                SELECT mt.Id, mt.Description AS Descr,
                       ISNULL(mt.Hardwired, 0) AS Hw,
                       COUNT(om.PeopleId) AS N
                FROM lookup.MemberType mt
                LEFT JOIN OrganizationMembers om ON om.MemberTypeId = mt.Id
                GROUP BY mt.Id, mt.Description, mt.Hardwired
                ORDER BY COUNT(om.PeopleId) DESC, mt.Description
            """):
                out.append({"id": int(r.Id),
                            "name": str(getattr(r, "Descr", "") or ""),
                            "hardwired": int(getattr(r, "Hw", 0) or 0),
                            "people": int(getattr(r, "N", 0) or 0)})
        except Exception as e:
            return safe_json({"success": False, "error": str(e)})
        return safe_json({"success": True, "types": out})

    if action == "list_roles":
        roles = []
        try:
            for r in q.QuerySql("SELECT RoleName FROM dbo.Roles ORDER BY RoleName"):
                nm = str(getattr(r, "RoleName", "") or "")
                if nm:
                    roles.append(nm)
        except Exception:
            pass
        return safe_json({"success": True, "roles": roles,
                          "can_share": can_edit_shared()})

    if action == "list_widgets":
        widgets, skipped = list_touchpoint_widgets()
        return safe_json({"success": True, "widgets": widgets,
                          "skipped": skipped})

    if action == "resolve_search":
        ids, note, err = _scope_ids(get_param("q_name", ""))
        if err:
            return safe_json({"success": False, "error": err})
        return safe_json({"success": True,
                          "ids": [int(i) for i in (ids or [])],
                          "note": note})

    if action == "custom_meta":
        doms = [d for d in custom_meta()
                if can_view_category(DOMAIN_CATEGORY.get(d.get("key", ""), ""))]
        return safe_json({"success": True, "domains": doms,
                          "filters": custom_filter_meta(),
                          "searches": list_saved_searches()})

    if action == "run_custom":
        raw = decode_payload(get_param("spec"))
        try:
            spec = json.loads(raw)
        except Exception as e:
            return safe_json({"success": False,
                              "error": "Bad tile spec: " + repr(e)})
        return safe_json(run_custom_tile(spec, get_param("global_range", ""),
                                        get_param("global_query", "")))

    if action == "drill_custom":
        raw = decode_payload(get_param("spec"))
        try:
            spec = json.loads(raw)
        except Exception as e:
            return safe_json({"success": False, "error": "Bad spec: " + repr(e)})
        return safe_json(drill_custom_tile(spec, decode_payload(get_param("d_label")),
                                          get_param("global_range", ""),
                                          get_param("global_query", "")))

    if action == "get_dashboard":
        act = acting_uid(uid)
        d = find_dashboard(act, get_param("d_id"), get_param("d_scope") or None)
        if not d:
            return safe_json({"success": False, "error": "Dashboard not found."})
        owner = ""
        if act != uid:
            for o in all_dashboard_owners():
                if o["uid"] == act:
                    owner = o["owner"]
        # What each tile can actually be narrowed by, worked out once here.
        # Offering a control a dashboard cannot use reads as broken, and the
        # only honest way to know is to look at the reports behind the tiles.
        for tab in (d.get("tabs") or []):
            for t in (tab.get("tiles") or []):
                cap = {"date": False, "program": False, "people": False}
                if t.get("kind") in ("links", "opscheck"):
                    # Not a report: it holds addresses, not data, so it is not
                    # counted for or against any bar control.
                    t["caps"] = None
                    continue
                if t.get("kind") == "custom":
                    # Built tiles always select people and always have dates.
                    cap = {"date": True, "program": False, "people": True}
                elif t.get("kind") == "report":
                    rep = catalog_report(t.get("report_id", ""))
                    if rep:
                        cap = report_caps(rep)
                    else:
                        # An Enterprise Reporting tile: its definition lives in
                        # the other script, so claim nothing rather than guess.
                        cap = {"date": True, "program": False, "people": False}
                t["caps"] = cap
        return safe_json({"success": True, "dashboard": d, "owner_name": owner,
                          "owner_uid": (act if act != uid else 0),
                          "can_edit": (d.get("scope") != "shared") or can_edit_shared()})

    if action == "add_from_library":
        lib_id = get_param("lib_id")
        # Reinstall: build the fresh copy FIRST, and only remove the old ones
        # once it exists. Deleting first would lose the dashboard outright if
        # the rebuild then failed.
        replace = get_param("replace", "") == "1"
        old_ids = []
        if replace:
            old_ids = [(d.get("id"), d.get("scope", "personal"))
                       for d in visible_dashboards(uid)
                       if d.get("from_library") == lib_id]
        ok, res = clone_from_library(uid, lib_id)
        if not ok:
            return safe_json({"success": False, "error": res})
        if replace:
            new_id = res.get("id") if isinstance(res, dict) else res
            for did, scope in old_ids:
                if str(did) != str(new_id):
                    delete_dashboard(uid, did, scope)
        if isinstance(res, dict):
            return safe_json({"success": True, "id": res.get("id"),
                              "missing": res.get("missing") or []})
        return safe_json({"success": True, "id": res})

    if action == "save_dashboard":
        raw = decode_payload(get_param("payload"))
        try:
            dash = json.loads(raw)
        except Exception as e:
            return safe_json({"success": False,
                              "error": "Could not parse dashboard: " + repr(e)})
        if not (dash.get("name") or "").strip():
            return safe_json({"success": False, "error": "Give the dashboard a name."})
        ok, res = store_dashboard(acting_uid(uid), dash)
        if not ok:
            return safe_json({"success": False, "error": res})
        return safe_json({"success": True, "id": res})

    if action == "delete_dashboard":
        ok, msg = delete_dashboard(acting_uid(uid), get_param("d_id"),
                                   get_param("d_scope", "personal"))
        return safe_json({"success": ok, "error": ("" if ok else msg)})

    if action == "new_dashboard":
        dash = {"id": "", "name": get_param("d_name", "New Dashboard"),
                "description": "", "icon": "fa-th-large", "roles": [],
                "scope": get_param("d_scope", "personal"),
                "cache_minutes": 30,
                "tabs": [{"name": "Tab 1", "tiles": []}]}
        ok, res = store_dashboard(uid, dash)
        if not ok:
            return safe_json({"success": False, "error": res})
        return safe_json({"success": True, "id": res})

    # Names itself. Both scripts used to answer with the same sentence, so a
    # request that went to the wrong one was indistinguishable from a stale
    # deploy of the right one.
    return safe_json({"success": False,
                      "error": "Unknown action '" + str(action)
                               + "' -- this was answered by TPxi_Dashboards, "
                                 "not by the reporting script."})


# ---------------------------------------------------------------------------
# ROUTER
# ---------------------------------------------------------------------------
_uid = current_user_id()

if model.HttpMethod == "post" and get_param("ajax") == "true":
    try:
        print handle_ajax(get_param("action", ""), _uid)
    except Exception:
        print safe_json({"success": False, "error": traceback.format_exc()[-800:]})

else:
    _rs = reports_script_name()
    # The shared library is used when it is installed, so there is one
    # implementation to fix. It is NOT required: almost no church has it, and
    # depending on it meant the update check silently did nothing -- which is
    # exactly what happened here, with 1.1.1 published and 1.1.0 running.
    update_js = ""
    try:
        _lib = model.PythonContent("TPxi_Lib_Update")
        if _lib:
            _ns = {}
            exec(_lib, _ns)
            update_js = _ns["update_check_js"](DC_SCRIPT_ID, APP_VERSION,
                                              apply_fn="applyAppUpdate")
    except Exception:
        update_js = ""
    if not update_js:
        update_js = inline_update_js()

    _t = {}

    def timed(label, fn):
        """Run fn, remember how long it took. Cheap enough to leave on."""
        t0 = time.time()
        try:
            return fn()
        finally:
            _t[label] = int((time.time() - t0) * 1000)

    boot = {
        "reports_script": _rs,
        "reports_ok": (1 if _rs else 0),
        "can_share": timed("can_share", can_edit_shared),
        "can_manage": timed("can_manage", can_manage_users),
        "searches": timed("searches", list_saved_searches),
        "programs": timed("programs", list_programs),
        "user_id": _uid,
        # Surfaced at boot rather than only inside Settings: someone whose
        # dashboards just vanished needs to be told where they went on the
        # screen they are already looking at.
        "legacy_count": timed(
            "legacy", lambda: len(legacy_personal().get("dashboards", []))),
        # For the links tile, which gates a shortcut on roles the way the
        # QuickLinks widget does. Display only: nothing behind a link is
        # protected by hiding it, and TouchPoint checks the target itself.
        "roles": timed("roles", my_roles),
        # Which reports the catalog can serve, so a tile saved before the
        # catalog existed still runs from it. The source used to be decided
        # once at install time and frozen into the tile, which left older
        # boards pointing at Enterprise Reporting for reports it never had.
        "catalog_ids": timed(
            "catalog", lambda: [str(r.get("id", "")) for r in
                                load_report_catalog().get("reports", [])]),
    }
    _t["total_ms"] = sum(v for k, v in _t.items() if k != "total_ms")
    boot["timing"] = _t

    page = []
    page.append('<link rel="stylesheet" href="' + GRIDSTACK_CSS + '" />')
    page.append('<link rel="stylesheet" href="' + GRIDSTACK_EXTRA_CSS + '" />')
    page.append('<' + 'script src="' + GRIDSTACK_JS + '"></' + 'script>')
    page.append('<' + 'script src="' + CHARTJS + '"></' + 'script>')
    page.append('<' + 'script src="' + GOOGLE_CHARTS + '"></' + 'script>')

    page.append('''
<style>
  /* TouchPoint prints its own "Dashboards" heading above us, 83px of it,
     saying exactly what our own title bar says one line lower. This rule
     only exists on this script's page, so nothing else in TouchPoint is
     affected. The empty .page-header strip above it goes too. */
  #page-header{display:none;}
  .box.box-responsive > .box-content{padding-top:0;}
  .db-bar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;
          padding:10px 0 14px 0;border-bottom:1px solid #e3e8ee;margin-bottom:12px;}
  .db-bar h2{margin:0;font-size:20px;font-weight:600;}
  .db-spacer{flex:1;}
  .db-tabs{display:flex;gap:2px;margin:0 0 12px 0;border-bottom:2px solid #e3e8ee;}
  .db-tab{padding:8px 16px;cursor:pointer;border:0;background:none;font-size:14px;
          color:#5b6875;border-bottom:2px solid transparent;margin-bottom:-2px;}
  .db-tab.on{color:#2f6fb5;border-bottom-color:#2f6fb5;font-weight:600;}
  .db-tile{background:#fff;border:1px solid #dde3ea;border-radius:6px;
           display:flex;flex-direction:column;overflow:hidden;height:100%;}
  .db-tile-hd{padding:8px 12px;border-bottom:1px solid #eef2f6;font-weight:600;
              font-size:13px;display:flex;align-items:center;gap:8px;
              background:#f8fafc;cursor:move;}
  .db-tile-bd{flex:1;padding:10px;overflow:auto;position:relative;min-height:60px;}
  .db-tile-x{margin-left:auto;color:#b0392f;cursor:pointer;font-weight:700;
             border:0;background:none;font-size:15px;line-height:1;}
  .db-muted{color:#8a94a0;font-size:12px;}
  .db-err{color:#a4232b;font-size:12px;}
  .db-card{border:1px solid #dde3ea;border-radius:6px;padding:14px;background:#fff;
           display:flex;flex-direction:column;gap:6px;}
  .db-card h4{margin:0;font-size:15px;}
  .db-grid-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));
                 gap:14px;}
  .db-modal{position:fixed;inset:0;background:rgba(15,23,32,.45);z-index:1050;
            display:none;align-items:flex-start;justify-content:center;padding:40px 16px;
            overflow:auto;}
  .db-modal.on{display:flex;}
  .db-modal-box{background:#fff;border-radius:8px;max-width:900px;width:100%;
                padding:20px;box-shadow:0 10px 40px rgba(0,0,0,.3);
                max-height:calc(100vh - 80px);display:flex;
                flex-direction:column;overflow:hidden;}
  /* The header stays put and the body takes the remaining height, scrolling
     inside the box. Without this a long table grew the box past the screen
     and painted over the page behind it. */
  #dbModalBody{flex:1;min-height:0;overflow:auto;}
  #dbModalBody .db-drill{max-height:none;}
  /* The table's own wrapper is the scroller. A sticky header sticks to its
     nearest scrollable ancestor, so if the body scrolled instead the header
     would slide away with the rows. Capped so the body itself does not also
     scroll, which would put two scrollbars on one list. */
  .db-tw{overflow:auto;}
  #dbModalBody .db-tw{max-height:calc(100vh - 260px);}
  .db-kpi{font-size:30px;font-weight:700;color:#2f6fb5;}
  table.db-t{width:100%;border-collapse:collapse;font-size:12px;}
  table.db-t th{text-align:left;padding:4px 6px;border-bottom:2px solid #e3e8ee;
                position:sticky;top:0;background:#fff;font-size:11px;
                text-transform:uppercase;color:#5b6875;}
  table.db-t td{padding:3px 6px;border-bottom:1px solid #f2f5f8;}
  .grid-stack-item-content{overflow:visible;}
  /* Tile bodies hold HTML written by other people: TouchPoint widgets, links
     tiles, church-authored dashboard widgets. An anchor or image is natively
     draggable, so if the pointer crosses one during a resize the browser
     starts its own drag, swallows the mouseup, and GridStack never gets the
     stop event. The tile then stays stuck to the cursor. */
  .db-tile-bd a, .db-tile-bd img{-webkit-user-drag:none;user-drag:none;}
  /* Belt and braces, and browser-agnostic: while the grid is being dragged or
     resized, nothing inside a tile can receive pointer events at all, so no
     content can start a competing gesture. */
  body.db-grabbing .db-tile-bd{pointer-events:none;user-select:none;}
  /* Sized by its own content, so it can report a SMALLER height than the tile
     it sits in. Measuring .db-tile-bd instead reports the tile's own height
     back to us and the tile can only ever grow. */
  .db-fit{display:block;}
  /* gridstack-extra covers 2 through 11 columns; one column is nobody's
     range, so spell it out or a phone shows an empty page. */
  .grid-stack.gs-1 > .grid-stack-item{width:100%;left:0;}
  .db-tile-fit{border:1px solid #cfd8e3;background:#fff;border-radius:3px;
               font-size:10px;line-height:1;padding:2px 5px;cursor:pointer;
               color:#5b6875;margin-left:auto;}
  .db-tile-fit.on{border-color:#9dc0e6;background:#eef4fb;color:#2f6fb5;}
  .db-chip{font-size:11px;padding:3px 9px;border:1px solid #d6dde4;
    border-radius:11px;background:#fff;color:#4a5561;cursor:pointer;
    text-transform:capitalize;}
  .db-chip:hover{border-color:#9fb3c8;}
  .db-chip.on{background:#26547c;border-color:#26547c;color:#fff;}
  .db-age{font-size:10px;color:#8a94a0;font-weight:normal;margin-left:6px;}
  .db-linkcat{font-size:11px;font-weight:600;color:#5b6875;
    text-transform:uppercase;letter-spacing:.03em;margin:8px 0 2px;
    padding-bottom:2px;border-bottom:1px solid #eef2f6;cursor:pointer;
    user-select:none;}
  .db-linkcat:hover{color:#2f6fb5;}
  .db-linkcat + .db-links{padding-left:10px;}
  .db-linkcat:first-child{margin-top:0;}
  .db-links{display:grid;gap:2px 0;padding:2px 0;text-align:center;
    grid-template-columns:repeat(auto-fill, minmax(80px, 1fr));}
  .db-link{min-width:80px;min-height:45px;height:auto;padding:2px 0;
    box-sizing:border-box;position:relative;display:flex;
    flex-direction:column;text-decoration:none;color:#005587;}
  .db-link:hover{text-decoration:none;color:#003f66;}
  .db-link:hover .db-link-ico i{color:#003f66;}
  .db-link-ico{display:flex;justify-content:center;align-items:center;
    width:100%;height:30px;margin:0;padding:0;}
  .db-link-ico i{color:#005587;font-size:24px;margin:0;
    transition:color 83ms linear;}
  .db-link-lbl{display:flex;justify-content:center;align-items:flex-start;
    width:100%;box-sizing:border-box;min-height:15px;height:auto;
    margin:0;padding:0 4px 2px 4px;overflow:visible;}
  .db-link-lbl span{font-size:11px;line-height:13px;margin:0;padding:0;
    overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;
    -webkit-box-orient:vertical;text-overflow:ellipsis;text-align:center;
    max-height:26px;width:100%;word-break:normal;word-wrap:break-word;}
  .db-ico{display:inline-flex;align-items:center;gap:5px;padding:4px 8px;
    border:1px solid #dde3ea;border-radius:4px;cursor:pointer;color:#5b6875;
    background:#fff;font-size:12px;}
  .db-ico:hover{border-color:#2f6fb5;color:#2f6fb5;}
  .db-ico2{display:flex;flex-direction:column;align-items:center;gap:3px;
    padding:6px 3px;border:1px solid transparent;border-radius:4px;
    cursor:pointer;font-size:10px;color:#5b6875;text-align:center;}
  .db-ico2 i{font-size:17px;color:#2f6fb5;}
  .db-ico2:hover{background:#eef4fb;}
  .db-ico2.on{border-color:#2f6fb5;background:#e3f2fd;}
  .db-roletag{display:inline-flex;align-items:center;gap:5px;font-size:11px;
    padding:2px 7px;border-radius:10px;background:#e7f1ff;color:#26547c;
    border:1px solid #cfe3ff;}
  .db-roletag span{cursor:pointer;font-weight:700;}
  .db-pinned{font-size:10px;font-weight:600;margin-left:6px;padding:1px 7px;
    border-radius:9px;background:#e7f1ff;color:#26547c;border:1px solid #cfe3ff;
    cursor:help;}
  .db-nofilter{font-size:10px;font-weight:600;margin-left:6px;padding:1px 7px;
    border-radius:9px;background:#fdf0d5;color:#8a5a00;border:1px solid #f0d9a8;
    cursor:help;}
  .db-tile-exp{margin-left:auto;border:0;background:none;cursor:pointer;
               color:#5b6875;font-size:13px;line-height:1;padding:2px 4px;}
  .db-tile-exp:hover{color:#2f6fb5;}
  .db-modal-wide .db-modal-box{max-width:1200px;}
  .db-drill{max-height:calc(100vh - 260px);min-height:220px;overflow:auto;}
  .db-actbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;
             padding:8px 10px;background:#f7f9fb;border:1px solid #dde3ea;
             border-radius:4px;
             position:sticky;top:0;z-index:3;}
  /* Headings stay put under the action bar. They need an opaque background,
     or rows show through as they scroll past. top:0 is relative to .db-tw,
     which already starts below the bar; the bar's higher z-index means they
     slide under it rather than leaving a strip rows can scroll through. */
  table.db-t thead th{position:sticky;top:0;z-index:2;background:#fff;
                      box-shadow:inset 0 -2px 0 #e3e8ee;}
  .db-people .db-tw{overflow-x:auto;}
  .db-age.live{color:#1b6b33;}
  .db-narrow{display:none;background:#eef4fb;border:1px solid #cfe0f2;
             color:#2f6fb5;border-radius:4px;padding:6px 10px;font-size:12px;
             margin-bottom:10px;}
  @media (max-width: 1100px){
    .db-bar h2{font-size:17px;}
    .db-bar select{min-width:140px;}
  }
  @media (max-width: 700px){
    .db-bar{gap:6px;}
    .db-bar select{width:100%;min-width:0;order:10;}
    .db-spacer{display:none;}
    .db-tabs{overflow-x:auto;}
    .db-tab{white-space:nowrap;}
  }
</style>
''')

    page.append('''
<div id="tpxiUpdate"></div>
<div class="db-bar">
  <h2 id="dbTitle">Dashboards</h2>
  <span id="dbScope" class="db-muted"></span>
  <div class="db-spacer"></div>
  <select id="dbPicker" style="padding:5px 8px;min-width:200px;"></select>
  <select id="dbRange" style="padding:5px 8px;display:none;"
          title="Date range applied to every tile that accepts one"
          onchange="dbSetRange()">
    <option value="">Each tile's own range</option>
    <option value="last_7_days">Last 7 days</option>
    <option value="last_30_days">Last 30 days</option>
    <option value="last_90_days">Last 90 days</option>
    <option value="last_6_months">Last 6 months</option>
    <option value="last_12_months">Last 12 months</option>
    <option value="last_24_months">Last 2 years</option>
    <option value="last_36_months">Last 3 years</option>
    <option value="last_48_months">Last 4 years</option>
    <option value="ytd">Year to date</option>
    <option value="fiscal_ytd">Fiscal year to date</option>
    <option value="last_calendar_year">Last calendar year</option>
    <option value="last_fiscal_year">Last fiscal year</option>
  </select>
  <select id="dbProgram" style="padding:5px 8px;display:none;max-width:180px;"
          title="Limit every tile that knows its program to one ministry area"
          onchange="dbSetProgram()">
    <option value="">All areas</option>
  </select>
  <select id="dbScopeQ" style="padding:5px 8px;display:none;max-width:190px;"
          title="Limit every tile that can be limited to a Search Builder search"
          onchange="dbSetQuery()">
    <option value="">Everyone</option>
  </select>
  <button class="btn btn-sm btn-default" id="dbRefreshBtn" onclick="dbRefresh()"
          style="display:none;" title="Discard cached tiles and fetch current data">Refresh</button>
  <button class="btn btn-sm btn-default" id="dbShareBtn" onclick="dbShare()"
          style="display:none;"
          title="The link to this dashboard, and who can open it">Share</button>
  <button class="btn btn-sm btn-default" onclick="dbSettings()"
          title="Church-wide settings for reports and dashboards">Settings</button>
  <button class="btn btn-sm btn-default" onclick="dbHelp()"
          title="How this dashboard counts things">?</button>
  <button class="btn btn-sm btn-default" onclick="dbLibrary()">Library<span
          id="libBadge" style="display:none;margin-left:6px;padding:1px 6px;
          border-radius:9px;background:#c0392b;color:#fff;font-size:11px;"></span></button>
  <button class="btn btn-sm btn-default" onclick="dbNew()">New</button>
  <button class="btn btn-sm btn-primary" id="dbEditBtn" onclick="dbToggleEdit()"
          style="display:none;">Edit</button>
</div>
<div id="dbTabs" class="db-tabs" style="display:none;"></div>
<div id="dbNarrowNote" class="db-narrow">Narrow screen: tiles are stacked to fit.
Open this on a wider screen to rearrange them.</div>
<div id="dbLegacyNote" style="display:none;border:1px solid #f0ad4e;
background:#fcf8e3;padding:8px 12px;border-radius:4px;margin-bottom:10px;"></div>
<div id="dbEditTools" style="display:none;margin-bottom:10px;">
  <button class="btn btn-xs btn-default" onclick="dbAddTile()">+ Add tile</button>
  <button class="btn btn-xs btn-default" onclick="dbAddTab()">+ Add tab</button>
  <button class="btn btn-xs btn-default" onclick="dbRenameTab()">Rename tab</button>
  <span style="margin-left:12px;font-size:12px;color:#5b6875;">Cache for
    <select id="dbCacheMins" onchange="dbSetCache()" style="padding:2px 4px;">
      <option value="0">off</option>
      <option value="5">5 min</option>
      <option value="15">15 min</option>
      <option value="30">30 min</option>
      <option value="60">1 hour</option>
      <option value="240">4 hours</option>
      <option value="1440">1 day</option>
    </select>
    <span class="db-muted">(tiles marked "live" always fetch fresh)</span>
  </span>
  <span style="margin-left:12px;font-size:12px;color:#5b6875;">Opens showing
    <select id="dbDefQuery" onchange="dbSetDefQuery()" style="padding:2px 4px;">
      <option value="">Everyone</option>
    </select>
    <span class="db-muted">(anyone can still change it for themselves)</span>
  </span>
  <button class="btn btn-xs btn-danger" style="float:right;" onclick="dbDelete()">Delete dashboard</button>
</div>
<div id="dbHome"></div>
<div class="grid-stack" id="dbGrid" style="display:none;"></div>

<div class="db-modal" id="dbModal"><div class="db-modal-box">
  <div style="display:flex;align-items:center;">
    <h3 id="dbModalTitle" style="margin:0;">Add</h3>
    <button class="btn btn-sm btn-default" style="margin-left:auto;"
            onclick="dbCloseModal()">Close</button>
  </div>
  <div id="dbModalBody" style="margin-top:14px;"></div>
</div></div>
''')

    page.append('<' + 'script>var DB_BOOT = ' + safe_json(boot) + ';</' + 'script>')

    page.append('''
<''' + '''script>
var DASH = null;          // dashboard being viewed
var TAB = 0;              // active tab index
var EDIT = false;
var GRID = null;

// Marks the page as mid-gesture so tile content cannot start a competing one.
// The listeners below are the safety net: if a stop event never arrives -- the
// exact failure this guards against -- the class would otherwise stay on and
// the whole dashboard would stop responding to clicks.
var GRABBING = false;

function gridGrabbing(on){
  GRABBING = !!on;
  try {
    if (on) document.body.classList.add("db-grabbing");
    else document.body.classList.remove("db-grabbing");
  } catch (e) { }
}

(function(){
  var clear = function(){ gridGrabbing(false); };
  document.addEventListener("mouseup", clear, true);
  document.addEventListener("pointerup", clear, true);
  document.addEventListener("dragend", clear, true);
  window.addEventListener("blur", clear);
})();
var CHARTS = {};          // tileKey -> Chart instance
var OBSERVERS = [];       // MutationObservers watching tiles that can grow
var AUTHORED_H = {};      // tile index -> the height the USER chose
var CELL_H = 60, CELL_MARGIN = 6, MAX_ROWS = 24;
var GRID_COLS = 12;       // columns currently displayed
var RESIZE_T = null;

function responsiveColumns(){
  var w = window.innerWidth || document.documentElement.clientWidth || 1200;
  if (w < 700) return 1;
  if (w < 1100) return 6;
  return 12;
}

function applyColumns(){
  if (!GRID) return;
  var c = responsiveColumns();
  if (c === GRID_COLS) return;
  GRID_COLS = c;
  // Narrow layouts are read-only, so packing beats preserving position:
  // moveScale keeps each tile's column, which leaves stair-step holes down
  // the page. compact keeps the authored ORDER and closes the gaps.
  var modes = ["compact", "list", "moveScale"];
  for (var m = 0; m < modes.length; m++){
    try { GRID.column(c, modes[m]); break; }
    catch (e) { if (m === modes.length - 1){ try { GRID.column(c); } catch (e2) { } } }
  }
  narrowGuard();
}

// Editing a layout that has been remapped to 6 or 1 columns would save those
// remapped coordinates over the real 12-column layout. So editing is only
// offered at full width, and any active edit is dropped on the way down.
function narrowGuard(){
  var narrow = GRID_COLS < 12;
  var note = $id("dbNarrowNote");
  if (note) note.style.display = (DASH && narrow) ? "block" : "none";
  var eb = $id("dbEditBtn"), tools = $id("dbEditTools");
  if (narrow && EDIT){
    EDIT = false;
    if (tools) tools.style.display = "none";
    // Through markClean so the button loses the "Save changes" styling too,
    // rather than sitting there green and saying Edit.
    markClean();
    renderTabs();
    renderGrid();
    return;
  }
  if (eb) eb.style.display = (DASH && DASH.can_edit && !narrow) ? "" : "none";
}

window.addEventListener("resize", function(){
  if (RESIZE_T) clearTimeout(RESIZE_T);
  RESIZE_T = setTimeout(applyColumns, 180);
});
var REPORTS = null;       // Enterprise Reporting catalog, loaded once
var REPORTS_ERR = "";     // why the report catalog is empty, if it is
var WIDGETS = null;           // TouchPoint DashboardWidgets, loaded once
var WIDGETS_SKIPPED = null;   // why some widgets are not offered
var WIDGETS_ERR = "";         // the list could not be fetched at all
var PICK_SRC = "reports"; // which source the Add-tile modal is showing
var CUSTOM_META = null;   // domains + saved searches for the tile builder
var CATALOG = null;       // published catalog: what is available and installed
var PICK_CAT = "";        // category chip currently selected in the picker
// Set only while an admin is working inside someone else's storage. Every
// request that touches storage carries it, so nothing can be written to the
// wrong person's slot by forgetting a parameter.
var ADMIN_OWNER = null;   // {uid, name} or null for "my own"
var ADMIN_CAN_DEL_SHARED = false;
var ROLES = null;         // role names available to limit a dashboard to

// The address bar is the share mechanism: a dashboard is reachable by link, and
// the link carries which tab you were on.
function readUrlState(){
  var out = {id: "", scope: "", tab: 0};
  try {
    var qs = (window.location.search || "").replace(/^\?/, "").split("&");
    for (var i = 0; i < qs.length; i++){
      var kv = qs[i].split("=");
      var k = decodeURIComponent(kv[0] || "");
      var v = decodeURIComponent((kv[1] || "").replace(/\+/g, " "));
      if (k === "d") out.id = v;
      else if (k === "s") out.scope = v;
      else if (k === "tab") out.tab = parseInt(v, 10) || 0;
    }
  } catch (e) { }
  return out;
}

function writeUrlState(){
  try {
    if (!window.history || !window.history.replaceState) return;
    var base = window.location.pathname;
    if (!DASH){ window.history.replaceState({}, "", base); return; }
    var u = base + "?d=" + encodeURIComponent(DASH.id)
          + "&s=" + encodeURIComponent(DASH.scope || "personal");
    if (TAB) u += "&tab=" + TAB;
    window.history.replaceState({}, "", u);
  } catch (e) { }
}

function $id(x){ return document.getElementById(x); }

// ---------------------------------------------------------------------------
// CACHE
// Kept in the browser, not on the server. Tile results are scoped to the person
// asking: their tags, their tasks, the searches they may run. A shared cache
// would hand one person another person's data, so this stays per-user and dies
// with the tab.
// ---------------------------------------------------------------------------
function nowMs(){ return new Date().getTime(); }

// One tile's payload may take this much of the store. The browser gives about
// 5 MB per origin in total, so this is deliberately well under it: a handful
// of large tiles must not be able to consume the whole budget between them.
var CACHE_MAX_BYTES = 1200000;

function cacheMinutes(){
  if (!DASH) return 0;
  if (DASH.cache_minutes === undefined || DASH.cache_minutes === null
      || DASH.cache_minutes === ""){
    // Never set on this dashboard: use the same default a new one gets,
    // rather than treating "unset" as "off".
    return 30;
  }
  var m = parseInt(DASH.cache_minutes, 10);
  if (isNaN(m) || m < 0) return 0;
  return m;
}

// What the cache is actually doing, per tile. Guessing at this from the
// outside is hopeless: a miss and a disabled cache look identical on screen.
function cacheReport(){
  var out = [], list = tiles();
  for (var i = 0; i < list.length; i++){
    var t = list[i], why = "", key = "", bytes = 0, age = -1;
    try { key = cacheKey(t); } catch (e) { why = "key failed: " + e.message; }
    if (!cacheMinutes()) why = "caching is off for this dashboard";
    else if (t.nocache) why = "tile is marked live";
    else if (key){
      var raw = null;
      try { raw = cacheStore().getItem(key); } catch (e) { why = "no storage"; }
      if (raw === null && !why){
        var mark = null;
        try { mark = cacheStore().getItem(key + ":toobig"); } catch (e) { }
        if (mark){
          var mb = 0;
          try { mb = (JSON.parse(mark) || {}).n || 0; } catch (e) { }
          why = "too big to cache (" + Math.round(mb / 1024) + " KB)";
        } else {
          why = "not stored";
        }
      }
      else if (raw){
        bytes = raw.length;
        try {
          var o = JSON.parse(raw);
          age = Math.floor((nowMs() - o.at) / 60000);
          why = (age > cacheMinutes()) ? ("expired, " + age + "m old") : "HIT";
        } catch (e) { why = "unreadable"; }
      }
    }
    out.push({title: t.title || t.report_id || t.check_id || t.kind,
              kind: t.kind || "report", why: why, bytes: bytes, key: key});
  }
  return out;
}

// What the page load actually spent its time on, and the slowest calls since.
// Reported rather than guessed at, because the person who sees it slow is not
// the person who can read the code.
function dbSpeedHtml(){
  var t = (DB_BOOT && DB_BOOT.timing) || {};
  var keys = [];
  for (var k in t){ if (k !== "total_ms") keys.push(k); }
  keys.sort(function(a, b){ return (t[b] || 0) - (t[a] || 0); });
  var h = "<h4 style='margin:14px 0 4px;'>Speed</h4>"
        + "<div>Page load spent <b>" + (t.total_ms || 0)
        + " ms</b> on the server:</div><ul style='margin:4px 0 8px 18px;'>";
  for (var i = 0; i < keys.length; i++){
    h += "<li>" + esc(keys[i]) + ": " + t[keys[i]] + " ms</li>";
  }
  h += "</ul>";
  var slow = CALL_LOG.slice(0).sort(function(a, b){ return b.ms - a.ms; }).slice(0, 8);
  if (slow.length){
    h += "<div>Slowest calls since this page opened:</div>"
      + "<table class='db-t' style='margin-top:4px;'><thead><tr><th>Action</th>"
      + "<th>ms</th><th>KB back</th></tr></thead><tbody>";
    for (var j = 0; j < slow.length; j++){
      h += "<tr><td>" + esc(slow[j].action) + "</td><td>" + slow[j].ms
        + "</td><td>" + Math.round(slow[j].bytes / 1024) + "</td></tr>";
    }
    h += "</tbody></table>";
  }
  return h;
}

function dbCacheDiag(){
  var box = $id("stCacheDiag");
  if (!box) return;
  var rows = cacheReport();
  var total = 0, hits = 0;
  for (var i = 0; i < rows.length; i++){
    total += rows[i].bytes;
    if (rows[i].why === "HIT") hits++;
  }
  var storeBytes = 0, storeKeys = 0;
  try {
    var st = cacheStore();
    for (var k = 0; k < st.length; k++){
      var kk = st.key(k);
      if (kk && kk.indexOf("tpxidb:") === 0){
        storeKeys++;
        storeBytes += (st.getItem(kk) || "").length;
      }
    }
  } catch (e) { }
  var h = "<div class='db-muted' style='margin-bottom:6px;'>This tab: <b>"
        + hits + "</b> of <b>" + rows.length + "</b> tiles cached, "
        + Math.round(total / 1024) + " KB. Whole browser store: " + storeKeys
        + " entries, " + Math.round(storeBytes / 1024) + " KB. Cache window: <b>"
        + cacheMinutes() + "</b> min.</div>"
        + "<table class='db-t'><thead><tr><th>Tile</th><th>Kind</th>"
        + "<th>State</th><th>KB</th></tr></thead><tbody>";
  for (var j = 0; j < rows.length; j++){
    h += "<tr><td>" + esc(rows[j].title) + "</td><td>" + esc(rows[j].kind)
      + "</td><td" + (rows[j].why === "HIT" ? " style='color:#1b6b33;'" : "")
      + ">" + esc(rows[j].why) + "</td><td>"
      + (rows[j].bytes ? Math.round(rows[j].bytes / 1024) : "") + "</td></tr>";
  }
  h += "</tbody></table>";
  box.innerHTML = h + dbSpeedHtml();
}

// localStorage, not sessionStorage: session storage is per TAB, so opening a
// dashboard in a new tab -- which is how a deep link or a fresh visit arrives
// -- started with an empty cache and refetched everything every time. The TTL
// still bounds staleness; this only changes how long the store survives.
function cacheStore(){
  try {
    var s = window.localStorage;
    s.setItem("tpxidb:probe", "1");
    s.removeItem("tpxidb:probe");
    return s;
  } catch (e) {
    // Private mode or a blocked origin: fall back rather than lose caching.
    try { return window.sessionStorage; } catch (e2) { return null; }
  }
}

// Cache entries are namespaced by user id. While current_user_id was broken
// every user was 0, so those entries are unreachable now that ids are real,
// and they would just sit there taking up the storage budget until eviction
// pushed out entries that are still good. Cleared once, on the first load
// after the fix.
function cachePurgeLegacy(){
  var st = cacheStore();
  if (!st) return 0;
  try {
    if (st.getItem("tpxidb:purged0") === "1") return 0;
  } catch (e) { return 0; }
  var dead = [];
  try {
    for (var i = 0; i < st.length; i++){
      var k = st.key(i);
      if (k && k.indexOf("tpxidb:0:") === 0) dead.push(k);
    }
    for (var j = 0; j < dead.length; j++){ st.removeItem(dead[j]); }
    st.setItem("tpxidb:purged0", "1");
  } catch (e) { }
  return dead.length;
}

function cacheKey(t){
  // Everything that changes the answer belongs in the key, or a tile edited to
  // point somewhere else would serve the old target's numbers.
  var base = "tpxidb:" + (DB_BOOT.user_id || 0) + ":" + ((DASH && DASH.id) || "") + ":";
  if (t.kind === "widget") return base + "w:" + t.widget_id;
  if (t.kind === "opscheck") return base + "o:" + t.check_id + ":" + (t.display || "");
  if (t.kind === "custom"){
    return base + "c:" + JSON.stringify(t.spec || {})
         + ":" + (t.own_dates ? "" : currentRange())
         + ":" + (t.own_scope ? "" : currentQuery());
  }
  return base + "r:" + tileSource(t) + ":" + t.report_id + ":" + (t.display || "")
       + ":" + (t.query || "") + ":" + (t.serving_types || "")
       + ":" + JSON.stringify(t.filters || {})
       + ":" + (t.own_dates ? "" : currentRange())
       + ":" + (t.own_scope ? "" : currentQuery())
       + ":" + (t.own_area ? "" : currentProgram())
       // The aggregation changes the answer as much as a filter does.
       + ":" + JSON.stringify(t.agg || {});
}

function cacheGet(t){
  if (!t || t.nocache) return null;
  var mins = cacheMinutes();
  if (!mins) return null;
  var store = cacheStore();
  if (!store) return null;
  try {
    var raw = store.getItem(cacheKey(t));
    if (!raw) return null;
    var o = JSON.parse(raw);
    if (!o || !o.at) return null;
    if ((nowMs() - o.at) > mins * 60000) return null;
    return o;
  } catch (e) { return null; }   // private mode, quota, or disabled storage
}

function cachePut(t, val){
  if (!t || t.nocache || !cacheMinutes()) return;
  var store = cacheStore();
  if (!store) return;
  var body;
  try {
    body = JSON.stringify({at: nowMs(), v: val});
  } catch (e) { return; }
  // Measured, not guessed: dh_members_missing_email returns 5,000 rows and
  // serialises to 666 KB here, so the old 500 KB ceiling silently refused to
  // cache it -- and every people-list tile on Data Health with it. Those were
  // exactly the tiles that "load every time". The ceiling now clears a result
  // of that size; anything genuinely enormous is still refused, but it SAYS
  // so instead of failing invisibly.
  if (body.length > CACHE_MAX_BYTES){
    // Drop the stale entry, then leave a marker so the diagnostic can explain
    // the miss. A result that outgrew the limit would otherwise keep serving
    // the previous, smaller answer as though it were current.
    try {
      store.removeItem(cacheKey(t));
      store.setItem(cacheKey(t) + ":toobig",
                    JSON.stringify({at: nowMs(), n: body.length}));
    } catch (e) { }
    return;
  }
  try {
    store.setItem(cacheKey(t), body);
  } catch (e) {
    // Out of room. Clearing only THIS dashboard was not enough: the store is
    // shared across every dashboard, so a board filled earlier in the session
    // keeps the quota and every later write fails silently -- which reads as
    // "the cache stopped working". Evict oldest-first across all of them,
    // then retry.
    if (!cacheEvictOldest(body.length)){ cacheClear(); }
    try { store.setItem(cacheKey(t), body); } catch (e2) { }
  }
}

// Free at least `need` bytes by dropping the least recently written entries.
// Returns false when there is nothing left to give up.
function cacheEvictOldest(need){
  var store = cacheStore();
  if (!store) return false;
  var mine = [];
  try {
    for (var i = 0; i < store.length; i++){
      var k = store.key(i);
      if (!k || k.indexOf("tpxidb:") !== 0) continue;
      var raw = store.getItem(k) || "";
      var at = 0;
      try { at = (JSON.parse(raw) || {}).at || 0; } catch (e) { at = 0; }
      mine.push({k: k, at: at, n: raw.length});
    }
  } catch (e) { return false; }
  if (!mine.length) return false;
  mine.sort(function(a, b){ return a.at - b.at; });
  var freed = 0;
  for (var j = 0; j < mine.length && freed < need * 2; j++){
    try { store.removeItem(mine[j].k); freed += mine[j].n; } catch (e) { }
  }
  return freed > 0;
}

// Editing a tile changes what it should show, so whatever is cached under its
// OLD settings has to go. The key is built from those settings, so this drops
// the pre-edit entry; the post-edit key simply has nothing stored yet.
function cacheDrop(t){
  if (!t) return;
  var store = cacheStore();
  if (!store) return;
  try { store.removeItem(cacheKey(t)); } catch (e) { }
}

function cacheClear(){
  var store = cacheStore();
  if (!store) return;
  try {
    var pre = "tpxidb:" + (DB_BOOT.user_id || 0) + ":" + ((DASH && DASH.id) || "") + ":";
    var kill = [];
    for (var i = 0; i < store.length; i++){
      var k = store.key(i);
      if (k && k.indexOf(pre) === 0) kill.push(k);
    }
    for (var j = 0; j < kill.length; j++){ store.removeItem(kill[j]); }
  } catch (e) { }
}

// Entries from other dashboards, or from an older visit, are still taking up
// the quota. Anything past its own dashboard's window is dead weight, so a
// load sweeps what has clearly expired.
function cacheSweep(){
  var store = cacheStore();
  if (!store) return;
  try {
    var cut = nowMs() - 24 * 60 * 60000;
    var kill = [];
    for (var i = 0; i < store.length; i++){
      var k = store.key(i);
      if (!k || k.indexOf("tpxidb:") !== 0) continue;
      try {
        var o = JSON.parse(store.getItem(k));
        if (!o || !o.at || o.at < cut) kill.push(k);
      } catch (e){ kill.push(k); }
    }
    for (var j = 0; j < kill.length; j++){ store.removeItem(kill[j]); }
  } catch (e) { }
}

// Marks a tile that cannot follow the dashboard's date range, so a number
// that refuses to move reads as "this report has no date dimension" rather
// than as a fault.
// What this tile pins for itself, shown in its header. Reads from the tile's
// own settings rather than a response, so it is right before the tile has
// finished loading.
function pinnedHtml(t){
  var bits = [];
  var f = t.filters || {};
  if (f.date_range) bits.push(rangeLabel(f.date_range));
  if (f.program || f.program_id){ bits.push(programLabel(f.program || f.program_id)); }
  if (t.query) bits.push(t.query);
  if (!bits.length) return "";
  return "<span class='db-pinned' title='Set on this tile, so the bar above "
       + "does not change it. Edit the tile to clear it.'>" + esc(bits.join(" | "))
       + "</span>";
}

function rangeLabel(v){
  var sel = $id("dbRange");
  if (sel){
    for (var i = 0; i < sel.options.length; i++){
      if (sel.options[i].value === v) return sel.options[i].text;
    }
  }
  return String(v).replace(/_/g, " ");
}

function programLabel(v){
  var ps = DB_BOOT.programs || [];
  for (var i = 0; i < ps.length; i++){
    if (String(ps[i].id) === String(v)) return ps[i].name;
  }
  return "one area";
}

// A report that depends on a church setting nobody has confirmed says so
// above its own numbers, with the way to fix it one click away.
var SETTING_LABELS = {
  bg_approval_ok: "which background check approvals count as cleared"
};

function needsSettingBar(box, r){
  var key = r && r.needs_setting;
  if (!key || !box) return;
  var bar = document.createElement("div");
  bar.style.cssText = "margin:0 0 8px;padding:6px 10px;border-radius:4px;"
    + "background:#fdf0d5;border:1px solid #f0d9a8;color:#8a5a00;font-size:12px;";
  bar.innerHTML = "This report needs someone to confirm "
    + esc(SETTING_LABELS[key] || key)
    + ". It is assuming the safest reading until then. "
    + "<a href='#' onclick='dbSettings();return false;'>Set it now</a>";
  box.insertBefore(bar, box.firstChild);
}

function scopeLabel(idx, r){
  var el = $id("scope" + idx);
  if (!el) return;
  if (!r){ el.textContent = ""; el.title = ""; return; }
  var bits = [], why = [];
  if (currentRange() && r.dated === false){
    bits.push("all dates");
    why.push("It has no date range to narrow, so the dashboard's date range "
             + "does not change it.");
  }
  if (currentQuery() && r.scopeable === false){
    bits.push("everyone");
    why.push("It does not count people one by one, so it cannot be limited to "
             + "a search.");
  }
  if (currentProgram() && r.has_program === false){
    bits.push("all areas");
    why.push("It does not record which program the number came from, so it "
             + "cannot be limited to one area.");
  }
  el.textContent = bits.join(", ");
  el.className = bits.length ? "db-nofilter" : "db-age";
  el.title = bits.length ? ("This tile is NOT narrowed by the bar. "
                            + why.join(" ")) : "";
}

function ageLabel(idx, at){
  var el = $id("age" + idx);
  if (!el) return;
  if (!at){ el.className = "db-age live"; el.textContent = ""; return; }
  var mins = Math.floor((nowMs() - at) / 60000);
  el.className = "db-age";
  el.textContent = mins < 1 ? "just now" : (mins + "m old");
  el.title = "Served from cache. Use Refresh for current numbers.";
}

// IronPython aside, the browsers this runs in predate Array.includes.
function contains(arr, v){
  for (var i = 0; i < (arr || []).length; i++){ if (arr[i] === v) return true; }
  return false;
}

function esc(s){
  return String(s === null || s === undefined ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ASP.NET rejects the whole POST if a value looks like markup, so angle
// brackets go over the wire as entities and the server puts them back.
function encPayload(s){
  return String(s).replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Installs a newer published build over this one. Named on window because the
// shared update banner looks the function up by name before it offers a button
// at all, rather than rendering one that throws on click.
window.applyAppUpdate = function(){
  if (!confirm("Install the newer version of this script over the one running "
               + "now? Dashboards, settings and installed reports are stored "
               + "separately and are not touched.")) return;
  var btn = document.getElementById("tpxiUpdateBtn");
  if (btn){ btn.disabled = true; btn.textContent = "Updating..."; }
  // The name the browser actually loaded, so a renamed install is written back
  // to the slot it came from.
  var nm = "";
  try {
    var m = (window.location.pathname || "").match(/\/PyScript(?:Form)?\/([^\/?#]+)/);
    if (m && m[1]) nm = m[1];
  } catch (e) { }
  post({action: "apply_update", script_name: nm}, function(r){
    if (!r || !r.success){
      alert((r && r.error) || "The update did not run.");
      if (btn){ btn.disabled = false; btn.textContent = "Update now"; }
      return;
    }
    alert("Updated " + (r.name || "this script") + ". Reloading.");
    window.location.reload();
  });
};

var CALL_LOG = [];      // {action, ms, bytes} for the slowest recent calls

function logCall(action, ms, bytes){
  CALL_LOG.push({action: action || "(none)", ms: ms, bytes: bytes});
  if (CALL_LOG.length > 60) CALL_LOG.shift();
}

function post(params, cb){
  var body = "ajax=true";
  for (var k in params){
    body += "&" + k + "=" + encodeURIComponent(params[k]);
  }
  var t0 = new Date().getTime();
  var xhr = new XMLHttpRequest();
  xhr.open("POST", window.location.pathname, true);
  xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
  xhr.onreadystatechange = function(){
    if (xhr.readyState !== 4) return;
    logCall(params.action, new Date().getTime() - t0,
            (xhr.responseText || "").length);
    var r = null;
    try { r = JSON.parse(xhr.responseText); }
    catch (e) {
      cb({success: false, error: "Bad response: " + xhr.responseText.substring(0, 300)});
      return;
    }
    cb(r);
  };
  xhr.send(body);
}

// Tiles get their data from Enterprise Reporting, not from this script.
function postReports(params, cb){
  var body = "ajax=true";
  for (var k in params){
    body += "&" + k + "=" + encodeURIComponent(params[k]);
  }
  var xhr = new XMLHttpRequest();
  xhr.open("POST", "/PyScriptForm/" + DB_BOOT.reports_script, true);
  xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
  // Add tile waits for THIS to come back before it opens. With no timeout a
  // slow or wedged reporting script meant the dialog simply never appeared,
  // with nothing on screen to say why.
  xhr.timeout = 15000;
  var done = false;
  var fail = function(why){
    if (done) return;
    done = true;
    cb({success: false, error: why});
  };
  xhr.ontimeout = function(){
    fail("The reporting script did not answer within 15 seconds.");
  };
  xhr.onerror = function(){ fail("Could not reach the reporting script."); };
  xhr.onreadystatechange = function(){
    if (xhr.readyState !== 4 || done) return;
    done = true;
    var r = null;
    try { r = JSON.parse(xhr.responseText); }
    catch (e) { r = {success: false, error: "Unreadable response from reports."}; }
    cb(r);
  };
  xhr.send(body);
}

// ---------- home ----------
function dbHome(){
  ADMIN_OWNER = null;
  DASH = null;
  writeUrlState();
  var sb0 = $id("dbShareBtn"); if (sb0) sb0.style.display = "none";
  $id("dbGrid").style.display = "none";
  $id("dbTabs").style.display = "none";
  $id("dbEditTools").style.display = "none";
  $id("dbEditBtn").style.display = "none";
  $id("dbTitle").textContent = "Dashboards";
  $id("dbScope").textContent = "";
  $id("dbHome").style.display = "";
  libUpdateBadge();
  post({action: "list_dashboards"}, function(r){
    if (!r.success){ $id("dbHome").innerHTML = "<p class='db-err'>" + esc(r.error) + "</p>"; return; }
    var h = "";
    if (!r.dashboards.length){
      h += "<p class='db-muted'>No dashboards yet. Start one from the "
        + "<b>Library</b>, or click <b>New</b> for a blank slate.</p>";
    }
    h += "<div class='db-grid-cards'>";
    for (var i = 0; i < r.dashboards.length; i++){
      var d = r.dashboards[i];
      h += "<div class='db-card'>"
        + "<h4><a href='#' onclick=\\"dbOpen('" + esc(d.id) + "','"
        + esc(d.scope) + "');return false;\\">"
        + esc(d.name) + "</a></h4>"
        + "<div class='db-muted'>" + esc(d.description || "") + "</div>"
        + "<div class='db-muted'>" + d.tabs + " tab(s) &middot; "
        + esc(d.scope) + "</div></div>";
    }
    h += "</div>";
    $id("dbHome").innerHTML = h;
    fillPicker(r.dashboards, "");
  });
}

// The picker used to be filled only while rendering the home view, so opening
// a dashboard any other way -- a deep link, straight out of the Library, or
// right after creating one -- left it blank.
var DASH_LIST = null;

function fillPicker(list, selectedId){
  if (list) DASH_LIST = list;
  var sel = $id("dbPicker");
  if (!sel || !DASH_LIST) return;
  var h = "<option value=''>All dashboards</option>";
  for (var j = 0; j < DASH_LIST.length; j++){
    h += "<option value='" + esc(DASH_LIST[j].id) + "'"
      + (String(DASH_LIST[j].id) === String(selectedId) ? " selected" : "")
      + ">" + esc(DASH_LIST[j].name) + "</option>";
  }
  sel.innerHTML = h;
}

// force=true after a create, install or delete, when the cached list is stale.
function loadPicker(selectedId, force){
  if (DASH_LIST && !force){ fillPicker(null, selectedId); return; }
  post({action: "list_dashboards"}, function(r){
    if (r && r.success) fillPicker(r.dashboards, selectedId);
  });
}

$id("dbPicker").onchange = function(){
  if (this.value) dbOpen(this.value);
  else dbHome();
};

// ---------- library ----------
// ---------------------------------------------------------------------------
// EXPAND
// A tile at full size, plus the rows behind it. Report tiles already carry
// their table: run_report returns grid_data alongside chart_data, so the data
// view costs nothing extra. Custom tiles carry their own rows. Widget tiles are
// rendered HTML with no structured data, so they only get the larger view.
// ---------------------------------------------------------------------------
var EXP_IDX = -1, EXP_MODE = "view";

function dbExpand(i){
  var t = tiles()[i];
  if (!t) return;
  EXP_IDX = i;
  EXP_MODE = "view";
  var p = LAST_PAYLOAD[i];
  if (p === undefined){
    dbModal(t.title || "Tile", "<p class='db-muted'>This tile has not finished "
          + "loading yet. Close this and try again in a moment.</p>");
    return;
  }
  showExpand();
}

function expandHasData(t){
  if (t.kind === "widget") return false;
  var p = LAST_PAYLOAD[EXP_IDX];
  if (!p) return false;
  if (t.kind === "custom") return !!(p.rows && p.rows.length);
  return !!(p.grid_data && (p.grid_data.rowData || p.grid_data.rows));
}

function showExpand(){
  var t = tiles()[EXP_IDX];
  if (!t) return;
  var p = LAST_PAYLOAD[EXP_IDX];
  var h = "<div class='db-tabs' style='margin-bottom:10px;'>"
        + "<button class='db-tab" + (EXP_MODE === "view" ? " on" : "")
        + "' onclick=\\"expandMode('view')\\">View</button>";
  if (expandHasData(t)){
    h += "<button class='db-tab" + (EXP_MODE === "data" ? " on" : "")
      + "' onclick=\\"expandMode('data')\\">Data</button>";
  }
  h += "</div><div id='expBody'></div>";
  document.getElementById("dbModal").className = "db-modal on db-modal-wide";
  $id("dbModalTitle").textContent = t.title || "Tile";
  $id("dbModalBody").innerHTML = h;
  drawExpand();
}

function expandMode(m){
  EXP_MODE = m;
  showExpand();
}

function drawExpand(){
  var t = tiles()[EXP_IDX];
  var p = LAST_PAYLOAD[EXP_IDX];
  var body = $id("expBody");
  if (!t || p === undefined || !body) return;

  if (EXP_MODE === "data"){
    // A catalog result is columns and rows, which the grid-shaped path does
    // not understand: it looked for grid_data, found none, and said "Nothing
    // to show". Expanded is also where the whole result belongs, so this view
    // is not capped the way the tile is.
    if (tileSource(t) === "catalog"){
      // Same shape the tile path uses: the state's box is the container and
      // the table is written into its first child.
      body.innerHTML = "<div class='db-drill'></div>";
      TABLE_STATE.exp = {cols: p.columns || [], rows: p.rows || [],
                         sort: (TABLE_STATE.exp || {}).sort,
                         box: body, tile: t, limit: 5000};
      drawCatalogTable("exp");
      return;
    }
    // A reporting-sourced grid gets the same treatment, so sorting and the
    // people actions do not depend on which source served the tile.
    if (t.kind === "report" && p && p.grid_data){
      var g2 = normalizeGrid(p.grid_data);
      body.innerHTML = "<div class='db-drill'></div>";
      TABLE_STATE.exp = {cols: g2.cols, rows: g2.rows,
                         sort: (TABLE_STATE.exp || {}).sort,
                         box: body, tile: t, limit: 5000};
      drawCatalogTable("exp");
      return;
    }
    body.innerHTML = "<div class='db-drill'>" + expandDataHtml(t, p) + "</div>";
    return;
  }
  if (t.kind === "widget"){
    // Widget markup is a page fragment that addresses its own elements by id.
    // Injecting a second copy here gives the document duplicate ids, so the
    // widget's own script finds the copy in the TILE behind this modal and
    // updates that one instead. The expanded copy then sits forever on its
    // loading message. An iframe gives it a document of its own.
    body.innerHTML = "<iframe src='/HomeWidgets/Embed/"
      + encodeURIComponent(t.widget_id)
      + "' style='width:100%;height:520px;border:0;' "
      + "title='" + esc(t.title || "Widget") + "'></iframe>";
    return;
  }
  // Charts need a sized parent because the canvas fills it; a table must NOT
  // get one, or it overflows the fixed height and spills out of the modal.
  var isChart = (t.kind === "custom")
              ? ((t.spec && t.spec.display) || "bar") !== "table"
              : (t.display === "chart");
  body.innerHTML = isChart
    ? "<div style='height:460px;'><div id='expCanvasHost' style='height:100%;'>"
      + "</div></div>"
    : "<div class='db-drill'><div id='expCanvasHost'></div></div>";
  var host = $id("expCanvasHost");
  if (t.kind === "custom"){
    renderCustom(host, "exp", p, (t.spec && t.spec.display) || "bar");
  } else if (tileSource(t) === "catalog"){
    renderCatalogResult(t, host, "exp", p);
  } else {
    renderReport(t, host, "exp", p);
  }
}

function expandDataHtml(t, p){
  if (t.kind === "custom"){
    var rows = p.rows || [];
    var canDrill = t.spec && t.spec.dimension && t.spec.dimension !== "none";
    var h = "<table class='db-t'><thead><tr><th>"
          + esc(p.dimension_label || "Group") + "</th><th style='text-align:right;'>"
          + esc(p.measure_label || "Value") + "</th>"
          + (canDrill ? "<th></th>" : "") + "</tr></thead><tbody>";
    for (var i = 0; i < rows.length; i++){
      h += "<tr><td>" + esc(rows[i].label) + "</td><td style='text-align:right;'>"
        + esc(fmtVal(rows[i].value, p.money)) + "</td>"
        + (canDrill ? ("<td style='text-align:right;'>"
            + "<button class='btn btn-xs btn-default' onclick='dbDrill(" + i
            + ")'>Who</button></td>") : "")
        + "</tr>";
    }
    h += "</tbody></table>";
    if (p.note) h += "<div class='db-muted'>" + esc(p.note) + "</div>";
    return h;
  }
  return gridTable(p.grid_data || {});
}

function dbDrill(rowIdx){
  var t = tiles()[EXP_IDX];
  var p = LAST_PAYLOAD[EXP_IDX];
  if (!t || !p || !p.rows || !p.rows[rowIdx]) return;
  var label = p.rows[rowIdx].label;
  var body = $id("expBody");
  body.innerHTML = "<span class='db-muted'>Looking up who is in "
                 + esc(label) + "...</span>";
  var dp = {action: "drill_custom", spec: encPayload(JSON.stringify(t.spec || {})),
            d_label: encPayload(label)};
  var gr2 = currentRange();
  if (gr2 && !t.own_dates){ dp.global_range = gr2; }
  var gq2 = currentQuery();
  if (gq2 && !t.own_scope){ dp.global_query = gq2; }
  post(dp, function(r){
    if (!r || !r.success){
      body.innerHTML = (r && r.restricted)
        ? restrictedHtml()
        : ("<span class='db-err'>" + esc((r && r.error) || "No response")
           + "</span>");
      return;
    }
    var ppl = r.people || [];
    DRILL_PEOPLE = ppl;
    var h = "<div style='margin-bottom:8px;'>"
          + "<button class='btn btn-xs btn-default' onclick='drawExpand()'>"
          + "&larr; Back</button> <b>" + esc(r.dimension_label || "") + ": "
          + esc(r.label) + "</b> <span class='db-muted'>" + ppl.length
          + " people</span></div>";
    if (!ppl.length){
      h += "<span class='db-muted'>Nobody matched.</span>";
    } else {
      h += "<div class='db-people'>" + peopleActionBar()
        + "<div class='db-drill'><table class='db-t'><thead><tr>"
        + "<th style='width:26px;'><input type='checkbox' class='pkAll' "
        + "onclick='pkToggleAll(this)'></th>"
        + "<th>Name</th><th>Age</th><th>Email</th></tr></thead><tbody>";
      for (var i = 0; i < ppl.length; i++){
        h += "<tr><td><input type='checkbox' class='pkOne' value='" + ppl[i].id
          + "' onclick='pkCount(this)'></td>"
          + "<td><a href='/Person2/" + ppl[i].id
          + "' target='_blank'>" + esc(ppl[i].name) + "</a></td><td>"
          + esc(ppl[i].age) + "</td><td>" + esc(ppl[i].email) + "</td></tr>";
      }
      h += "</tbody></table></div></div>";
    }
    if (r.capped){
      h += "<div class='db-muted'>Showing the first 500.</div>";
    }
    if (r.note){ h += "<div class='db-muted'>" + esc(r.note) + "</div>"; }
    body.innerHTML = h;
  });
}

// ---------------------------------------------------------------------------
// BULK ACTIONS on a drilled list
// ---------------------------------------------------------------------------
var DRILL_PEOPLE = [];

// Expanding a tile puts a SECOND people list in the document while the tile's
// own list is still there. These were ids, so every lookup found the tile's
// copy first: ticking a box in the modal updated the bar behind it and the
// modal's own buttons never enabled. Everything here is therefore scoped to
// the .db-people block the click came from.
function personHref(id, payload){
  var tab = (payload && payload.display && payload.display.person_tab) || "";
  // Only a plain fragment name is accepted: this ends up in an href, and the
  // value travels in from the published catalog.
  if (!/^[a-z0-9_-]+$/i.test(tab)) tab = "";
  return "/Person2/" + encodeURIComponent(id) + (tab ? ("#" + tab) : "");
}

function peopleActionBar(){
  return "<div class='db-actbar'>"
    + "<span class='pkCountLbl db-muted'>None selected</span>"
    + "<span style='flex:1;'></span>"
    + "<button class='btn btn-xs btn-default pkTagBtn' disabled "
    + "onclick='pkTag(this)'>Tag selected</button>"
    + "<button class='btn btn-xs btn-default pkTaskBtn' disabled "
    + "onclick='pkTask(this)'>Add task</button></div>";
}

// The block a control belongs to. Falls back to the document so a caller
// without an element still behaves as it always did.
function pkRoot(el){
  var e = el;
  while (e && e.className !== undefined){
    if (String(e.className).indexOf("db-people") >= 0) return e;
    e = e.parentNode;
  }
  return document;
}

// A table only draws part of what it loaded, so the drawn checkboxes are not
// the selection. Where a block names a table state, that state holds it.
function pkStateFor(root){
  if (!root || !root.getAttribute) return null;
  var key = root.getAttribute("data-sel");
  if (key === null || key === undefined || key === "") return null;
  return TABLE_STATE[key] || TABLE_STATE[parseInt(key, 10)] || null;
}

function pkSelected(root){
  var st = pkStateFor(root);
  if (st && st.selected){
    var out = [];
    for (var k in st.selected){ if (st.selected[k]) out.push(k); }
    return out;
  }
  var boxes = (root || document).querySelectorAll(".pkOne");
  var got = [];
  for (var i = 0; i < boxes.length; i++){
    if (boxes[i].checked) got.push(boxes[i].value);
  }
  return got;
}

// One row ticked or unticked.
function pkPick(cb){
  var root = pkRoot(cb);
  var st = pkStateFor(root);
  if (st){
    st.selected = st.selected || {};
    if (cb.checked){ st.selected[cb.value] = 1; }
    else { delete st.selected[cb.value]; }
  }
  pkCount(cb);
}

function pkCount(el){
  var root = pkRoot(el);
  var sel = pkSelected(root);
  var lbl = root.querySelector(".pkCountLbl");
  if (lbl){
    lbl.textContent = sel.length ? (sel.length + " selected") : "None selected";
  }
  var tb = root.querySelector(".pkTagBtn"), kb = root.querySelector(".pkTaskBtn");
  if (tb) tb.disabled = !sel.length;
  if (kb) kb.disabled = !sel.length;
}

function pkToggleAll(cb){
  var root = pkRoot(cb);
  var st = pkStateFor(root);
  if (st){
    // Every LOADED row, not the 200 on screen. Ticking the header box and
    // getting 200 of 1,295 was a quiet way to act on the wrong people.
    st.selected = {};
    if (cb.checked){
      var pid = peopleCol(st.cols || []);
      for (var r = 0; r < (st.rows || []).length; r++){
        st.selected[String(st.rows[r][pid])] = 1;
      }
    }
  }
  var boxes = root.querySelectorAll(".pkOne");
  for (var i = 0; i < boxes.length; i++){ boxes[i].checked = !!cb.checked; }
  pkCount(cb);
}

function pkMsg(root, t){
  var el = (root || document).querySelector(".pkCountLbl");
  if (el) el.textContent = t;
}

function pkTag(btn){
  var root = pkRoot(btn);
  var sel = pkSelected(root);
  if (!sel.length) return;
  var name = prompt("Tag " + sel.length + " people with which tag?", "");
  if (!name || !name.trim()) return;
  pkMsg(root, "Tagging " + sel.length + "...");
  // Handled by this script, so it works on a tile whose report came from the
  // catalog with no other script installed.
  post({action: "bulk_tag", people_ids: sel.join(","),
        tag_name: name.trim()}, function(r){
    if (!r || !r.success){
      alert((r && r.error) || "Tagging failed.");
      pkCount(btn);
      return;
    }
    pkMsg(root, r.message || ("Tagged " + sel.length + "."));
  });
}

function pkTask(btn){
  var sel = pkSelected(pkRoot(btn));
  if (!sel.length) return;
  // The task dialog reuses the one modal, so opening it from an expanded tile
  // replaces that view. Remembered here and restored on close, otherwise
  // adding a task silently throws away the list you picked from.
  TASK_FROM_EXPAND = ($id("expBody") !== null && $id("expBody") !== undefined);
  var h = "<p class='db-muted'>A task will be created for each of the "
        + sel.length + " selected people.</p>"
        + "<div><label>What needs doing</label>"
        + "<textarea id='tkMsg' style='width:100%;height:90px;padding:6px;' "
        + "placeholder='Follow up about...'></textarea></div>"
        + "<div style='display:grid;grid-template-columns:1fr 1fr;gap:12px;"
        + "margin-top:10px;'>"
        + "<div><label>Assign to</label>"
        + "<input id='tkWho' style='width:100%;padding:5px;' "
        + "placeholder='Type a name, or leave blank for yourself' "
        + "oninput='tkSearch()'>"
        + "<div id='tkWhoList' class='db-muted' style='font-size:12px;'></div>"
        + "<div id='tkWhoPick' class='db-muted' style='font-size:12px;'></div></div>"
        + "<div><label>Due date (optional)</label>"
        + "<input id='tkDue' style='width:100%;padding:5px;' "
        + "placeholder='mm/dd/yyyy'></div></div>"
        + "<div style='margin-top:12px;'><label>Keywords</label>"
        + "<input id='tkKwFind' style='width:100%;padding:5px;' "
        + "placeholder='Filter keywords...' oninput='tkKwFilter()'>"
        + "<div id='tkKwList' style='max-height:150px;overflow:auto;"
        + "border:1px solid #ddd;padding:6px;margin-top:4px;'>"
        + "<span class='db-muted'>Loading keywords...</span></div>"
        + "<div id='tkKwNote' class='db-muted' style='margin-top:4px;'></div>"
        + "<div id='tkQs' style='margin-top:8px;'></div>"
        + "</div>"
        + "<div style='margin-top:12px;'>"
        + "<button class='btn btn-sm btn-primary' onclick='pkTaskSave()'>"
        + "Create tasks</button> <span id='tkMsgOut' class='db-muted'></span></div>";
  dbModal("Add a task for " + sel.length + " people", h);
  TASK_PIDS = sel;
  TASK_ASSIGNEE = "";
  loadTaskKeywords();
}

var KEYWORDS = null;

function loadTaskKeywords(){
  if (KEYWORDS){ drawTaskKeywords(); return; }
  post({action: "task_keywords"}, function(r){
    KEYWORDS = (r && r.success) ? (r.keywords || []) : [];
    if (!KEYWORDS.length){
      var el = $id("tkKwList");
      if (el){
        el.innerHTML = "<span class='db-muted'>"
          + esc((r && r.error) || "No active keywords.") + "</span>";
      }
      return;
    }
    drawTaskKeywords();
  });
}

function drawTaskKeywords(){
  var box = $id("tkKwList");
  if (!box) return;
  var h = "";
  for (var i = 0; i < KEYWORDS.length; i++){
    var k = KEYWORDS[i];
    h += "<label class='tkKwRow' data-name='"
      + esc((k.name + " " + k.code).toLowerCase())
      + "' style='display:block;font-weight:normal;margin:0 0 3px;'>"
      + "<input type='checkbox' class='tkKw' value='" + k.id
      + "' onclick='tkKwPicked()'> " + esc(k.name)
      // Flagged because the answers cannot be filled in from here: whoever
      // works the task supplies them on the task itself.
      + (k.questions
         ? (" <span style='font-size:11px;color:#8a5a00;'>"
            + k.questions + " question" + (k.questions === 1 ? "" : "s")
            + "</span>") : "")
      + "</label>";
  }
  box.innerHTML = h;
}

function tkKwFilter(){
  var v = (($id("tkKwFind") || {}).value || "").toLowerCase();
  var rows = document.querySelectorAll("#tkKwList .tkKwRow");
  for (var i = 0; i < rows.length; i++){
    var nm = rows[i].getAttribute("data-name") || "";
    rows[i].style.display = (!v || nm.indexOf(v) >= 0) ? "block" : "none";
  }
}

function tkKwSelected(){
  var out = [];
  var boxes = document.querySelectorAll(".tkKw");
  for (var i = 0; i < boxes.length; i++){
    if (boxes[i].checked) out.push(boxes[i].value);
  }
  return out;
}

function tkKwPicked(){
  var sel = tkKwSelected(), anyQ = false;
  for (var i = 0; i < (KEYWORDS || []).length; i++){
    for (var j = 0; j < sel.length; j++){
      if (String(KEYWORDS[i].id) === sel[j] && KEYWORDS[i].questions) anyQ = true;
    }
  }
  var note = $id("tkKwNote"), host = $id("tkQs");
  if (note) note.innerHTML = "";
  if (!host) return;
  if (!anyQ || !sel.length){
    host.innerHTML = "";
    return;
  }
  host.innerHTML = "<span class='db-muted'>Loading questions...</span>";
  post({action: "keyword_questions", keyword_ids: sel.join(",")}, function(r){
    if (!r || !r.success){
      host.innerHTML = "<span class='db-err'>"
        + esc((r && r.error) || "Could not load the questions.") + "</span>";
      return;
    }
    drawKeywordQuestions(r.questions || []);
  });
}

// 1 and 2 are a heading and an instruction block in TouchPoint's own task
// form, so they are shown as text rather than asked as questions.
function drawKeywordQuestions(qs){
  var host = $id("tkQs");
  if (!qs.length){ host.innerHTML = ""; return; }
  var h = "<div style='border:1px solid #dde3ea;background:#f7f9fb;padding:10px;'>"
        + "<div style='font-weight:600;margin-bottom:6px;'>Keyword questions</div>";
  var last = "";
  for (var i = 0; i < qs.length; i++){
    var qq = qs[i], id = "kq_" + qq.id;
    if (qq.keyword !== last){
      h += "<div class='db-muted' style='margin:6px 0 2px;'>" + esc(qq.keyword)
        + "</div>";
      last = qq.keyword;
    }
    if (qq.type === 1){
      h += "<div style='font-weight:600;'>" + esc(qq.name) + "</div>";
      continue;
    }
    if (qq.type === 2){
      h += "<div class='db-muted'>" + esc(qq.name) + "</div>";
      continue;
    }
    h += "<div style='margin-bottom:6px;'><label style='font-weight:normal;'>"
      + esc(qq.name) + "</label>";
    if (qq.type === 4){
      h += "<textarea id='" + id + "' class='kq' data-q='" + qq.id
        + "' data-t='" + qq.type + "' style='width:100%;height:60px;padding:5px;'>"
        + "</textarea>";
    } else if (qq.type === 5){
      h += "<select id='" + id + "' class='kq' data-q='" + qq.id
        + "' data-t='" + qq.type + "' style='width:100%;padding:5px;'>"
        + "<option value=''>-- not answered --</option>";
      for (var o = 0; o < (qq.options || []).length; o++){
        h += "<option value='" + qq.options[o].id + "'>"
          + esc(qq.options[o].name) + "</option>";
      }
      h += "</select>";
    } else if (qq.type === 6){
      h += "<select id='" + id + "' class='kq' data-q='" + qq.id
        + "' data-t='" + qq.type + "' style='width:100%;padding:5px;'>"
        + "<option value=''>-- not answered --</option>"
        + "<option value='true'>Yes</option><option value='false'>No</option>"
        + "</select>";
    } else if (qq.type === 7){
      h += "<div id='" + id + "' class='kq' data-q='" + qq.id
        + "' data-t='" + qq.type + "'>";
      for (var c = 0; c < (qq.options || []).length; c++){
        h += "<label style='display:block;font-weight:normal;'>"
          + "<input type='checkbox' class='kqcb' value='" + qq.options[c].id
          + "'> " + esc(qq.options[c].name) + "</label>";
      }
      h += "</div>";
    } else {
      // 3 = single line, 8 = date
      h += "<input id='" + id + "' class='kq' data-q='" + qq.id
        + "' data-t='" + qq.type + "' style='width:100%;padding:5px;'"
        + (qq.type === 8 ? " placeholder='mm/dd/yyyy'" : "") + ">";
    }
    h += "</div>";
  }
  h += "<div class='db-muted'>Left blank is fine. Whoever works the task can "
    + "still answer these on the task itself.</div></div>";
  host.innerHTML = h;
}

function tkAnswers(){
  var out = {}, fields = document.querySelectorAll("#tkQs .kq");
  for (var i = 0; i < fields.length; i++){
    var f = fields[i], qid = f.getAttribute("data-q");
    var t = parseInt(f.getAttribute("data-t"), 10);
    if (t === 7){
      var picked = [], boxes = f.querySelectorAll(".kqcb");
      for (var b = 0; b < boxes.length; b++){
        if (boxes[b].checked) picked.push(boxes[b].value);
      }
      if (picked.length) out[qid] = picked;
    } else if (f.value !== undefined && String(f.value).length){
      out[qid] = f.value;
    }
  }
  return out;
}

var TASK_PIDS = [], TASK_ASSIGNEE = "", TK_T = null;
var TASK_FROM_EXPAND = false;

function tkSearch(){
  var v = ($id("tkWho") || {}).value || "";
  if (TK_T) clearTimeout(TK_T);
  if (v.length < 2){ $id("tkWhoList").innerHTML = ""; return; }
  TK_T = setTimeout(function(){
    post({action: "search_assignee", search_term: v}, function(r){
      if (!r || !r.success){ return; }
      var h = "";
      for (var i = 0; i < (r.people || []).length; i++){
        var pp = r.people[i];
        h += "<div><a href='#' onclick=\\"tkPick(" + pp.id + ",'"
          + esc(pp.name).replace(/'/g, "") + "');return false;\\">"
          + esc(pp.name) + "</a></div>";
      }
      $id("tkWhoList").innerHTML = h || "<i>no match</i>";
    });
  }, 250);
}

function tkPick(id, name){
  TASK_ASSIGNEE = String(id);
  $id("tkWhoPick").innerHTML = "Assigning to <b>" + esc(name) + "</b>";
  $id("tkWhoList").innerHTML = "";
}

function pkTaskSave(){
  var msg = (($id("tkMsg") || {}).value || "").trim();
  if (!msg){ alert("Say what the task is."); return; }
  var params = {action: "bulk_task", people_ids: TASK_PIDS.join(","),
                task_message: msg};
  if (TASK_ASSIGNEE) params.assignee_id = TASK_ASSIGNEE;
  var kw = tkKwSelected();
  if (kw.length) params.keyword_ids = kw.join(",");
  var ans = tkAnswers();
  for (var k in ans){
    params.answers = encPayload(JSON.stringify(ans));
    break;
  }
  var due = (($id("tkDue") || {}).value || "").trim();
  if (due) params.due_date = due;
  $id("tkMsgOut").textContent = "Creating...";
  post(params, function(r){
    if (!r || !r.success){
      $id("tkMsgOut").textContent = "";
      alert((r && r.error) || "Could not create tasks.");
      return;
    }
    // Say how many actually landed: bulk_task swallows per-person failures,
    // so "created 12" out of 15 selected is information worth having.
    $id("tkMsgOut").textContent = (r.message || "Done.")
      + (r.count !== undefined && r.count !== TASK_PIDS.length
         ? (" (" + TASK_PIDS.length + " selected)") : "");
  });
}

function dbCopyLink(){
  var box = $id("dbLinkBox");
  if (!box) return;
  try {
    box.select();
    document.execCommand("copy");
    $id("dbCopyMsg").textContent = "Copied.";
  } catch (e) {
    $id("dbCopyMsg").textContent = "Press Ctrl+C to copy.";
  }
}

// ---------------------------------------------------------------------------
// SHARING
// Roles here are enforced server-side in user_can_see and store_dashboard, not
// merely hidden in this dialog. Nothing below is the security boundary.
// ---------------------------------------------------------------------------
function dbShare(){
  if (!DASH) return;
  if (ROLES){ showShare(); return; }
  post({action: "list_roles"}, function(r){
    ROLES = (r && r.success) ? r.roles : [];
    showShare();
  });
}

function showShare(){
  var scope = DASH.scope || "personal";
  var chosen = DASH.roles || [];
  var u = window.location.protocol + "//" + window.location.host
        + window.location.pathname + "?d=" + encodeURIComponent(DASH.id)
        + "&s=" + encodeURIComponent(scope) + (TAB ? ("&tab=" + TAB) : "");
  var h = "<div style='margin-bottom:14px;'><label>Link to this dashboard</label>"
        + "<div style='display:flex;gap:8px;'>"
        + "<input id='dbLinkBox' style='flex:1;padding:6px;' value='"
        + esc(u) + "' readonly>"
        + "<button class='btn btn-sm btn-primary' onclick='dbCopyLink()'>Copy"
        + "</button></div>"
        + "<div id='dbCopyMsg' class='db-muted'></div>"
        // Said here rather than in a separate dialog: whether a link works for
        // anyone else is the same question as who the dashboard is shared with.
        + (scope !== "shared"
           ? ("<div class='db-muted'>This is a <b>personal</b> dashboard, so "
              + "the link only works for you. Set it to shared below to let "
              + "others open it.</div>")
           : (chosen.length
              ? ("<div class='db-muted'>Only people holding: "
                 + esc(chosen.join(", ")) + "</div>")
              : "<div class='db-muted'>Anyone who can open Dashboards can view "
                + "this.</div>"))
        + "</div>";
  h += "<div style='display:grid;grid-template-columns:1fr 1fr;gap:12px;'>";
  h += "<div><label>Name</label>"
     + "<input id='shName' style='width:100%;padding:5px;' value='"
     + esc(DASH.name || "") + "'></div>";
  h += "<div><label>Description</label>"
     + "<input id='shDesc' style='width:100%;padding:5px;' value='"
     + esc(DASH.description || "") + "'></div>";
  h += "</div>";

  h += "<div style='margin-top:12px;'><label>Who owns it</label><div>";
  h += "<label style='font-weight:normal;'><input type='radio' name='shScope' "
     + "value='personal'" + (scope === "personal" ? " checked" : "")
     + "> Just me</label> &nbsp; ";
  if (DB_BOOT.can_share){
    h += "<label style='font-weight:normal;'><input type='radio' name='shScope' "
       + "value='shared'" + (scope === "shared" ? " checked" : "")
       + "> Shared with staff</label>";
  } else {
    h += "<span class='db-muted'>(You do not have permission to create shared "
       + "dashboards.)</span>";
  }
  h += "</div></div>";

  h += "<div style='margin-top:12px;'><label>Who can see it</label>"
     + "<div class='db-muted'>Tick nothing and everyone who can open Dashboards "
     + "can view it. Tick one or more roles to limit it. This applies to shared "
     + "dashboards; a personal one is only ever visible to you.</div>"
     + "<div style='max-height:220px;overflow:auto;border:1px solid #dde3ea;"
     + "border-radius:4px;padding:8px;margin-top:6px;display:grid;"
     + "grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:4px;'>";
  for (var i = 0; i < (ROLES || []).length; i++){
    var rn = ROLES[i];
    var on = false;
    for (var j = 0; j < chosen.length; j++){ if (chosen[j] === rn) on = true; }
    h += "<label style='font-weight:normal;font-size:12px;'>"
      + "<input type='checkbox' class='shRole' value='" + esc(rn) + "'"
      + (on ? " checked" : "") + "> " + esc(rn) + "</label>";
  }
  h += "</div></div>";
  h += "<div style='margin-top:14px;'>"
     + "<button class='btn btn-sm btn-primary' onclick='dbShareSave()'>Save</button>"
     + " <span id='shMsg' class='db-muted'></span></div>";
  if (!(DASH.can_edit)){
    // Viewer: the link is theirs to copy, the settings are not theirs to change.
    h = h.substring(0, h.indexOf("<div style='display:grid;grid-template-columns:1fr 1fr;gap:12px;'>"))
      + "<div class='db-muted'>You can open and share this dashboard, but not "
      + "change who it belongs to.</div>";
  }
  dbModal("Share this dashboard", h);
}

function dbShareSave(){
  var name = ($id("shName") || {}).value || "";
  if (!name.trim()){ alert("Give the dashboard a name."); return; }
  var scope = DASH.scope || "personal";
  var radios = document.querySelectorAll("input[name='shScope']");
  for (var i = 0; i < radios.length; i++){
    if (radios[i].checked) scope = radios[i].value;
  }
  var roles = [];
  var boxes = document.querySelectorAll(".shRole");
  for (var j = 0; j < boxes.length; j++){
    if (boxes[j].checked) roles.push(boxes[j].value);
  }
  var prev = DASH.scope || "personal";
  DASH.name = name.trim();
  DASH.description = ($id("shDesc") || {}).value || "";
  DASH.roles = roles;
  DASH.scope = scope;
  if (prev !== scope){ DASH.prev_scope = prev; }
  harvestLayout();
  post({action: "save_dashboard", payload: encPayload(JSON.stringify(DASH))},
       function(r){
    if (!r.success){ alert(r.error); DASH.scope = prev; return; }
    DASH.id = r.id;
    delete DASH.prev_scope;
    dbCloseModal();
    // Moving stores changes the id, so reopen rather than leaving a stale one
    // in the address bar.
    dbOpen(DASH.id, DASH.scope, TAB);
  });
}

var ADMIN_ROWS = [];      // every dashboard the admin screen is showing

function adminLoad(){
  var box = $id("stAdminList");
  if (!box) return;
  box.innerHTML = "<span class='db-muted'>Loading...</span>";
  post({action: "admin_dashboards"}, function(r){
    if (!r || !r.success){
      box.innerHTML = "<span class='db-err'>" + esc((r && r.error) || "Failed")
                    + "</span>";
      return;
    }
    ADMIN_ROWS = r.rows || [];
    ADMIN_CAN_DEL_SHARED = !!r.can_delete_shared;
    LEGACY = r.legacy || [];
    var pre = legacyPanelHtml();
    if (!ADMIN_ROWS.length){
      box.innerHTML = pre
        + "<span class='db-muted'>No dashboards yet.</span>";
      return;
    }
    box.innerHTML = pre
      + "<input id='adminFind' placeholder='Filter by person or dashboard name...' "
      + "style='width:100%;padding:6px 10px;margin-bottom:8px;' "
      + "oninput='adminRender()'>"
      // Reassigning is the same job as claiming an orphan, so it uses the same
      // picker and the same tick-and-go shape rather than a second idiom.
      + "<div style='margin-bottom:8px;padding:8px;background:#f6f8fa;"
      + "border-radius:3px;display:flex;gap:8px;align-items:flex-start;'>"
      + "<div style='flex:1;'>"
      + pickerHtml("rabulk", "Move the ticked ones to...") + "</div>"
      + "<button class='btn btn-primary btn-sm' onclick='adminReassign()'>"
      + "Move ticked</button></div>"
      + "<div id='adminRows' style='max-height:340px;overflow:auto;'></div>";
    adminRender();
  });
}

var LEGACY = [];      // dashboards saved before per-user storage worked

// ---------------------------------------------------------------------------
// PERSON PICKER
// One picker serves both tables. Name alone is not enough: someone with two
// records has the same name on both, so age, login and email are shown, which
// is what actually tells them apart.
// ---------------------------------------------------------------------------
var PICK = {};        // picker id -> {id: PeopleId, name: string}
var PICK_T = {};      // picker id -> pending search timer

function pickerHtml(pid, placeholder){
  return "<input id='pk_" + pid + "' placeholder='" + esc(placeholder || "Type a name...")
    + "' style='width:100%;padding:4px 6px;' autocomplete='off' "
    + "data-pk='" + esc(pid) + "' oninput='pkFind(this)'>"
    + "<div id='pkh_" + pid + "' style='font-size:12px;'></div>";
}

function pkFind(el){
  var pid = el.getAttribute("data-pk");
  if (PICK_T[pid]) clearTimeout(PICK_T[pid]);
  PICK[pid] = null;              // typing again clears any earlier choice
  PICK_T[pid] = setTimeout(function(){
    var v = (el.value || "").trim();
    var hits = $id("pkh_" + pid);
    if (!hits) return;
    if (v.length < 2){ hits.innerHTML = ""; return; }
    hits.innerHTML = "<span class='db-muted'>Looking...</span>";
    post({action: "search_assignee", search_term: v}, function(r){
      var ppl = (r && r.people) || [];
      if (!ppl.length){
        hits.innerHTML = "<span class='db-muted'>Nobody with a login matches "
                       + "that.</span>";
        return;
      }
      var h = "";
      for (var i = 0; i < ppl.length; i++){
        var pp = ppl[i];
        // Everything that separates one account from another sits on the row,
        // because two records for the same person share a name AND an age.
        var extra = [];
        if (pp.age >= 0) extra.push("age " + pp.age);
        if (pp.user) extra.push(esc(pp.user));
        if (pp.email) extra.push(esc(pp.email));
        extra.push("id " + pp.id);
        h += "<div style='padding:2px 0;'><a href='#' data-pk='" + esc(pid)
          + "' data-pid='" + pp.id + "' data-nm='" + esc(pp.name)
          + "' onclick='pkChoose(this);return false;'>" + esc(pp.name)
          + "</a> <span class='db-muted'>" + extra.join(" &middot; ")
          + "</span></div>";
      }
      hits.innerHTML = h;
    });
  }, 250);
}

function pkChoose(el){
  var pid = el.getAttribute("data-pk");
  var who = parseInt(el.getAttribute("data-pid"), 10);
  var nm = el.getAttribute("data-nm") || "";
  PICK[pid] = {id: who, name: nm};
  var box = $id("pk_" + pid);
  if (box) box.value = nm;
  var hits = $id("pkh_" + pid);
  if (hits) hits.innerHTML = "Will go to <b>" + esc(nm) + "</b> (id " + who + ")";
}

function pkGet(pid){
  return (PICK[pid] && PICK[pid].id) || 0;
}

// ---------------------------------------------------------------------------
// CLAIMING THE QUARANTINED SET
// ---------------------------------------------------------------------------

function legacyPanelHtml(){
  if (!LEGACY.length) return "";
  var h = "<div style='border:1px solid #f0ad4e;background:#fcf8e3;padding:10px 12px;"
        + "border-radius:4px;margin-bottom:12px;'>"
        + "<b>" + LEGACY.length + " dashboard(s) are waiting to be claimed.</b>"
        + "<div style='margin-top:6px;'>An earlier version could not tell users "
        + "apart, so every personal dashboard was written to one shared place "
        + "and everyone could see everyone else's. They are now hidden from "
        + "everybody until someone says who each one belongs to. Nothing was "
        + "deleted.</div>"
        + "<div style='margin:10px 0;padding:8px;background:#fff;border-radius:3px;'>"
        + "<div style='display:flex;gap:8px;align-items:flex-start;'>"
        + "<div style='flex:1;'>" + pickerHtml("lgbulk", "Give the ticked ones to...")
        + "</div>"
        + "<button class='btn btn-primary btn-sm' onclick='lgAssignChecked()'>"
        + "Assign ticked</button>"
        + "<button class='btn btn-default btn-sm' onclick='lgDropChecked()'>"
        + "Delete ticked</button></div></div>"
        + "<table class='db-t' style='background:#fff;'>"
        + "<thead><tr>"
        + "<th style='width:24px;'><input type='checkbox' onclick='lgAll(this)'></th>"
        + "<th>Dashboard</th><th>Size</th><th>Probably built by</th>"
        + "</tr></thead><tbody>";
  for (var i = 0; i < LEGACY.length; i++){
    var d = LEGACY[i];
    // The hint comes from the Search Builder searches its tiles use, since
    // dbo.Query records who owns a search. It is evidence, not proof.
    var who = "", first = null;
    var seen = {};
    for (var j = 0; j < (d.hints || []).length; j++){
      var hn = d.hints[j];
      if (!hn.owner || seen[hn.owner]) continue;
      seen[hn.owner] = 1;
      who += (who ? ", " : "") + esc(hn.owner);
      if (!first && hn.uid) first = hn;
    }
    if (!who) who = "<span class='db-muted'>no clue in the data</span>";
    h += "<tr><td><input type='checkbox' class='lgck' value='" + d.idx + "'></td>"
      + "<td>" + esc(d.name) + "</td>"
      + "<td class='db-muted'>" + d.tabs + " tab(s), " + d.tiles + " tile(s)</td>"
      + "<td>" + who
      + (first ? (" <a href='#' data-pk='lgbulk' data-pid='" + first.uid
                  + "' data-nm='" + esc(first.owner)
                  + "' onclick='pkChoose(this);return false;'>use</a>") : "")
      + "</td></tr>";
  }
  return h + "</tbody></table></div>";
}

function lgAll(box){
  var cks = document.querySelectorAll(".lgck");
  for (var i = 0; i < cks.length; i++){ cks[i].checked = box.checked; }
}

function lgChecked(){
  var out = [], cks = document.querySelectorAll(".lgck");
  for (var i = 0; i < cks.length; i++){
    if (cks[i].checked) out.push(cks[i].value);
  }
  return out;
}

function lgAssignChecked(){
  var picked = lgChecked();
  if (!picked.length){ alert("Tick the ones to move first."); return; }
  var to = pkGet("lgbulk");
  if (!to){ alert("Choose who they belong to first."); return; }
  post({action: "legacy_assign", idxs: picked.join(","), to_uid: to}, function(r){
    if (!r || !r.success){ alert((r && r.error) || "Failed"); return; }
    DASH_LIST = null;
    PICK = {};
    adminLoad();
  });
}

// Deleted highest index first, since removing one shifts the ones after it.
function lgDropChecked(){
  var picked = lgChecked();
  if (!picked.length){ alert("Tick the ones to delete first."); return; }
  if (!confirm("Permanently delete " + picked.length + " dashboard(s)?")) return;
  picked.sort(function(a, b){ return Number(b) - Number(a); });
  var i = 0;
  function nextOne(){
    if (i >= picked.length){ PICK = {}; adminLoad(); return; }
    post({action: "legacy_drop", idx: picked[i++]}, function(r){
      if (!r || !r.success){ alert((r && r.error) || "Failed"); adminLoad(); return; }
      nextOne();
    });
  }
  nextOne();
}

function adminRender(){
  var box = $id("adminRows");
  if (!box) return;
  var v = (($id("adminFind") || {}).value || "").toLowerCase();
  var h = "<table class='db-t'><thead><tr>"
        + "<th style='width:24px;'><input type='checkbox' onclick='adminAll(this)'></th>"
        + "<th>Person</th><th>Dashboard</th>"
        + "<th>Where</th><th>Tabs</th><th>Tiles</th><th></th></tr></thead><tbody>";
  var shown = 0, lastOwner = "";
  for (var i = 0; i < ADMIN_ROWS.length; i++){
    var d = ADMIN_ROWS[i];
    var hay = ((d.owner || "") + " " + (d.name || "")).toLowerCase();
    if (v && hay.indexOf(v) < 0) continue;
    shown++;
    var isShared = d.scope === "shared";
    // Dimmed rather than blanked on a repeat: the grouping still reads, and
    // the row you tick always says whose it is. Tracked against what is
    // actually DRAWN, since the filter can hide the first row of a group.
    var repeat = (d.owner === lastOwner);
    var who = esc(d.owner || "");
    lastOwner = d.owner;
    h += "<tr><td><input type='checkbox' class='rack' value='"
      + esc(d.uid + ":" + d.scope + ":" + d.id) + "'></td>"
      + "<td" + (repeat ? " style='color:#8a97a5;'" : "") + ">" + who
      + (who && d.mine ? " <span class='db-muted'>(you)</span>" : "")
      + "</td><td>" + esc(d.name || "(unnamed)") + "</td><td>"
      + (isShared
         ? "<span style='font-size:11px;padding:1px 7px;border-radius:9px;"
           + "background:#e7f1ff;color:#26547c;'>shared</span>"
           + (d.roles ? " <span class='db-muted' style='font-size:11px;'>"
                        + esc(d.roles) + "</span>" : "")
         : "<span class='db-muted' style='font-size:11px;'>personal</span>")
      + "</td><td>" + d.tabs + "</td><td>" + d.tiles + "</td><td style='white-space:nowrap;'>"
      // Values on data attributes: a person or dashboard name can contain an
      // apostrophe, which would end an inlined JS string literal.
      + "<button class='btn btn-xs btn-default' data-i='" + i
      + "' onclick='adminOpen(this)'>Open</button> "
      + ((isShared && !ADMIN_CAN_DEL_SHARED)
         ? ""
         : ("<button class='btn btn-xs btn-danger' data-i='" + i
            + "' onclick='adminDelete(this)'>Delete</button>"))
      + "</td></tr>";
  }
  h += "</tbody></table>";
  if (!shown){ h += "<span class='db-muted'>Nothing matches that.</span>"; }
  box.innerHTML = h;
}

function adminAll(box){
  var cks = document.querySelectorAll(".rack");
  for (var i = 0; i < cks.length; i++){ cks[i].checked = box.checked; }
}

function adminReassign(){
  var refs = [], cks = document.querySelectorAll(".rack");
  for (var i = 0; i < cks.length; i++){
    if (cks[i].checked) refs.push(cks[i].value);
  }
  if (!refs.length){ alert("Tick the dashboards to move first."); return; }
  var to = pkGet("rabulk");
  if (!to){ alert("Choose who they should belong to first."); return; }
  var anyShared = false;
  for (var j = 0; j < refs.length; j++){
    if (refs[j].split(":")[1] === "shared") anyShared = true;
  }
  var note = anyShared
    ? " A shared one stays shared; only who owns it changes."
    : "";
  if (!confirm("Move " + refs.length + " dashboard(s) to "
               + PICK["rabulk"].name + "?" + note)) return;
  post({action: "reassign", refs: refs.join(","), to_uid: to}, function(r){
    if (!r || !r.success){ alert((r && r.error) || "Failed"); return; }
    DASH_LIST = null;
    PICK = {};
    adminLoad();
  });
}

function adminOpen(btn){
  var d = ADMIN_ROWS[parseInt(btn.getAttribute("data-i"), 10)];
  if (!d) return;
  // Shared dashboards are not in anyone's slot, and neither is your own board,
  // so neither needs the admin context.
  ADMIN_OWNER = (d.scope === "shared" || d.mine)
              ? null : {uid: d.uid, name: d.owner};
  dbCloseModal();
  dbOpen(d.id, d.scope, 0);
}

function adminDelete(btn){
  var d = ADMIN_ROWS[parseInt(btn.getAttribute("data-i"), 10)];
  if (!d) return;
  var whose = d.scope === "shared" ? "the shared dashboards"
                                   : (d.mine ? "your dashboards" : d.owner);
  if (!confirm("Delete " + (d.name || "this dashboard") + " from " + whose
               + "? This cannot be undone.")) return;
  var pp = {action: "delete_dashboard", d_id: d.id, d_scope: d.scope};
  if (d.scope !== "shared" && !d.mine){ pp.owner_uid = d.uid; }
  post(pp, function(r){
    if (!r || !r.success){ alert((r && r.error) || "Delete failed."); return; }
    DASH_LIST = null;
    adminLoad();
  });
}

// Written to answer the questions this tool actually provokes: why a number
// did not move, and why it disagrees with a report someone already trusts.
function dbHelp(){
  var h = "";

  h += "<h4 style='margin:0 0 4px;'>The three controls in the bar</h4>"
    + "<p class='db-muted' style='margin-top:0;'>Each one narrows every tile "
    + "that can honor it, and leaves the rest alone. A tile that cannot be "
    + "narrowed says so with an amber tag beside its title, so a number that "
    + "refuses to move is never a mystery.</p>"
    + "<ul>"
    + "<li><b>Date range</b> applies to any report with a date filter. A tile "
    + "tagged <i>all dates</i> has none, usually because it is a current-state "
    + "list such as a pending queue or an account roster.</li>"
    + "<li><b>Area</b> limits to one Program. A tile tagged <i>all areas</i> "
    + "does not record which program its number came from.</li>"
    + "<li><b>Search</b> limits to a Search Builder search. A tile tagged "
    + "<i>everyone</i> counts meetings rather than people, so there is no "
    + "person to match against. Head-counted worship is the common case: the "
    + "meeting knows 1,243 attended, not who they were.</li>"
    + "</ul>"
    + "<p class='db-muted'>Between area and search almost everything narrows: "
    + "charts by area, name lists by search. A tile you set a filter on "
    + "yourself always keeps its own, whatever the bar says.</p>";

  h += "<h4 style='margin:14px 0 4px;'>How attendance is counted</h4>"
    + "<ul>"
    + "<li><b>Head counts are included.</b> Where a meeting has a head count "
    + "larger than the number of individual check-ins, the head count is used. "
    + "Without that most of Sunday morning is invisible: over the last year "
    + "the Worship program logged 4,780 individual check-ins against 177,325 "
    + "actually counted, and across every involvement it is 201,533 against "
    + "487,803 -- more than half of all attendance.</li>"
    + "<li><b>A day is a day, not a service.</b> Involvements that meet three "
    + "times on a Sunday show one Sunday. <i>AvgPerDay</i> is the day total; "
    + "<i>AvgPerMeeting</i> is the per-service figure beside it.</li>"
    + "<li><b>Week At A Glance decides what counts</b>, unless you change it "
    + "in Settings. See below.</li>"
    + "</ul>";

  h += "<h4 style='margin:14px 0 4px;'>Week At A Glance</h4>"
    + "<p style='margin-top:0;'>TouchPoint's own Week At A Glance report shows "
    + "a Program when it has a <b>Report Group</b>, and a Division as a row "
    + "under it when that Division has a <b>Report Line</b> (the number sets "
    + "the display order, so having one is what includes it). Set both under "
    + "<i>Administration &gt; Organizations &gt; Divisions</i>."
    + " <a href='https://docs.touchpointsoftware.com/SummaryReports/"
    + "WeekAtAGlanceSpecs.html' target='_blank' rel='noopener'>"
    + "TouchPoint's specification</a>.</p>"
    + "<p>By default this dashboard counts the same involvements, so its "
    + "numbers line up with the report your staff already read. Switching "
    + "<b>Settings &gt; Weekly attendance counts</b> to every meeting adds "
    + "everything deliberately left off it, which here is choir and orchestra "
    + "rehearsals, Peer Place, ESL and the enrichment classes. Both are "
    + "legitimate; they answer different questions, and the gap is roughly "
    + "750 a week.</p>";

  h += "<h4 style='margin:14px 0 4px;'>Fresh numbers</h4>"
    + "<p style='margin-top:0;'>Tiles are cached for the period set in edit "
    + "mode, so a dashboard left open does not re-query on every visit. "
    + "<b>Refresh</b> discards the cache and fetches now. A tile marked "
    + "<i>live</i> always fetches. Changing a bar control or a church setting "
    + "clears the cache by itself, because the old answer was to a different "
    + "question.</p>";

  h += "<h4 style='margin:14px 0 4px;'>Where tiles come from</h4>"
    + "<p style='margin-top:0;'><b>Library</b> holds ready-made dashboards and "
    + "the report catalog behind them. A red count on the button means "
    + "published definitions have changed since you installed them; refreshing "
    + "a copy leaves your layout alone. Report help, including what each one "
    + "counts and which church settings it depends on, is shown when you add "
    + "it as a tile.</p>";

  h += "<h4 style='margin:14px 0 4px;'>Who can see what</h4>"
    + "<p style='margin-top:0;'>Personal dashboards are yours. Shared ones are "
    + "limited by role. Data access is checked when a tile runs, not when it "
    + "is placed, so a giving tile on a dashboard someone shared with you still "
    + "shows nothing unless you are on the finance team.</p>";

  dbModal("About this dashboard", h);
}

function dbSettings(){
  post({action: "get_settings"}, function(r){
    if (!r || !r.success){ alert("Could not read settings."); return; }
    var detected = r.reports_script || "";

    // What reports MEAN comes first. The reporting-script name and the file
    // installer are setup-once plumbing and were sitting above the values
    // people actually adjust.
    var h = "<h4 style='margin:0 0 6px;'>Values your reports use</h4>"
          + "<p class='db-muted' style='margin-top:0;'>These change what the "
          + "numbers mean. A tile that depends on one says so when you add it.</p>";

    var months = ["January", "February", "March", "April", "May", "June",
                  "July", "August", "September", "October", "November",
                  "December"];
    h += "<div style='margin-bottom:10px;'>"
       + "<label>Background check approvals you accept</label>"
       + "<div id='stApprovals' style='display:flex;flex-wrap:wrap;gap:12px;"
       + "padding:6px 0;'>";
    var appr = r.bg_approvals || [], on = r.bg_approval_ok || [];
    for (var ai = 0; ai < appr.length; ai++){
      var isOn = false;
      for (var aj = 0; aj < on.length; aj++){
        if (on[aj] === appr[ai]) isOn = true;
      }
      h += "<label style='font-weight:normal;'><input type='checkbox' "
        + "class='stAppr' value='" + esc(appr[ai]) + "'"
        + (isOn ? " checked" : "") + "> " + esc(appr[ai]) + "</label>";
    }
    h += "</div><div class='db-muted'>Which values in a background check's "
       + "<b>Approval</b> box mean the person is cleared. Anything not ticked "
       + "counts as no valid check, so someone serving with minors on it is "
       + "reported."
       + (r.bg_approval_set ? "" : " <b>Not set yet</b>, so reports are "
          + "assuming Approved only.")
       + "</div></div>";

    h += "<div style='margin-bottom:10px;'><label>Weekly attendance counts</label>"
       + "<select id='stAttScope' style='width:100%;padding:6px;'>"
       + "<option value='waag'"
       + (r.attendance_scope !== "all" ? " selected" : "")
       + ">Week At A Glance involvements</option>"
       + "<option value='all'"
       + (r.attendance_scope === "all" ? " selected" : "")
       + ">Every meeting in TouchPoint</option></select>"
       + "<div class='db-muted'>Week At A Glance counts only divisions you gave "
       + "a report line, so matching it keeps this dashboard and that report in "
       + "step. Every meeting adds rehearsals, enrichment and anything else "
       + "deliberately left off it.</div></div>";

    h += "<div><label>Fiscal year starts in</label>"
       + "<select id='stFiscal' style='width:100%;padding:6px;'>";
    for (var mi = 0; mi < 12; mi++){
      h += "<option value='" + (mi + 1) + "'"
        + ((mi + 1) === (r.fiscal_month || 10) ? " selected" : "") + ">"
        + months[mi] + "</option>";
    }
    // Wrong here and "this fiscal year" quietly means the wrong twelve months.
    h += "</select><div class='db-muted'>Used by the fiscal-year date filters "
       + "on tiles.</div></div>";

    h += "<div style='margin-top:12px;'>"
       + "<label>A background check stays valid for</label>"
       + "<input id='stBgDays' style='width:100%;padding:6px;' value='"
       + esc(String(r.bg_check_days || 730)) + "'>"
       + "<div class='db-muted'>Days. Used by the safe-church reports to decide "
       + "whether a check has expired.</div></div>";

    if (!r.can_share){
      h += "<div class='db-err' style='margin-top:10px;'>You do not have "
         + "permission to change these.</div>";
    } else {
      h += "<div style='margin-top:12px;'>"
         + "<button class='btn btn-sm btn-primary' onclick='dbSettingsSave()'>"
         + "Save</button> <span id='stMsg' class='db-muted'></span></div>";
    }

    h += "<h4 style='margin:16px 0 6px;'>Is the cache working?</h4>"
       + "<p class='db-muted' style='margin-top:0;'>Open the dashboard you are "
       + "asking about first, then come here. A cache miss and a cache that is "
       + "switched off look identical on screen, so this says which it is, per "
       + "tile.</p>"
       + "<button class='btn btn-sm btn-default' onclick='dbCacheDiag()'>"
       + "Check the tile cache</button>"
       + "<div id='stCacheDiag' style='margin-top:8px;'></div>";

    h += "<div style='margin-top:16px;border-top:1px solid #e3e8ee;"
       + "padding-top:10px;'>"
       + "<a href='#' id='stAdvToggle' onclick='stToggleAdvanced();return false;'>"
       + "Show advanced</a></div>";

    if (DB_BOOT.can_manage){
      h += "<h4 style='margin:16px 0 6px;'>Everyone's dashboards</h4>"
        + "<p class='db-muted' style='margin-top:0;'>Personal dashboards belong "
        + "to the person who made them. Opening one here lets you fix or remove "
        + "it for them, and the bar says whose it is the whole time.</p>"
        + "<button class='btn btn-sm btn-default' onclick='adminLoad()'>"
        + "Show everyone's dashboards</button>"
        + "<div id='stAdminList' style='margin-top:8px;'></div>";
    }
    h += "<div id='stAdvanced' style='display:none;margin-top:10px;'>";
    h += "<div><label>Enterprise Reporting script name</label>"
       + "<input id='stName' style='width:100%;padding:6px;' value='"
       + esc(r.configured || "") + "' placeholder='"
       + (detected ? esc(detected) : "e.g. TPxi_EnterpriseReporting") + "'>"
       + "<div class='db-muted'>Currently using: <b>"
       + (detected ? esc(detected) : "none found") + "</b>"
       + (r.configured ? " (set here)" : " (detected)") + ". "
       + "Only needed for reports built inside that script. Catalog reports, "
       + "widgets and built tiles all work without it.</div></div>";
    h += "<div style='margin-top:12px;'>"
       + "<button class='btn btn-sm btn-default' onclick='dbDiagnose()'>"
       + "Diagnose report tiles</button>"
       + "<div class='db-muted'>Shows which stored script TouchPoint will "
       + "actually run, and whether it is the one you edited.</div>"
       + "<div id='stDiag' style='margin-top:8px;'></div></div>";
    h += "<h4 style='margin:14px 0 4px;'>Deploy a script into TouchPoint</h4>"
       + "<div class='db-muted'>Writes a .py file straight into Special "
       + "Content, bypassing the editor. This is how any TPxi script gets "
       + "installed when a paste will not commit: the body is entity-encoded "
       + "and sent in parts, so ASP.NET does not reject it as markup or "
       + "exceed the request size. Not limited to this dashboard.</div>"
       + "<div style='display:flex;gap:8px;align-items:center;margin-top:8px;"
       + "flex-wrap:wrap;'>"
       + "<input id='upName' style='padding:5px;width:230px;' "
       + "placeholder='Script name, e.g. EnterpriseReporting'>"
       + "<input type='file' id='upFile' accept='.py,.txt'>"
       + "<button class='btn btn-sm btn-primary' onclick='dbUploadPick()'>"
       + "Upload and write</button></div>"
       + "<div id='upMsg' class='db-muted' style='margin-top:6px;'></div>";
    h += "</div>";

    dbModal("Settings", h);
    // Orphaned dashboards are the reason someone is most likely here, so the
    // list opens itself rather than hiding behind another button.
    if (DB_BOOT.can_manage && (DB_BOOT.legacy_count || 0) > 0) adminLoad();
  });
}

// Advanced is collapsed rather than removed: the file installer is how a
// script gets into TouchPoint at all when the editor refuses a large paste,
// and the reporting-script name still matters to a church with its own
// reports built inside Enterprise Reporting.
function stToggleAdvanced(){
  var box = $id("stAdvanced"), link = $id("stAdvToggle");
  if (!box) return;
  var open = box.style.display !== "none";
  box.style.display = open ? "none" : "";
  if (link){
    link.textContent = open
      ? "Show advanced (reporting script, diagnostics, install a script from a file)"
      : "Hide advanced";
  }
}

function dbUploadPick(){
  var f = $id("upFile");
  if (!f || !f.files || !f.files.length){ alert("Choose a .py file first."); return; }
  var file = f.files[0];
  var name = (($id("upName") || {}).value || "").trim();
  if (!name){ alert("Enter the script name to write to."); return; }
  var msg = $id("upMsg");
  msg.textContent = "Reading " + file.name + "...";
  var fr = new FileReader();
  fr.onload = function(){
    var text = String(fr.result || "");
    if (text.length < 1000){ msg.textContent = "That file looks too small."; return; }
    uploadScript(name, text, msg);
  };
  fr.onerror = function(){ msg.textContent = "Could not read the file."; };
  fr.readAsText(file);
}

// Ampersand FIRST, then the angle brackets: encoding < before & would turn a
// literal &lt; in the source into &amp;lt; on the way back and corrupt it.
// The server decodes in the reverse order for the same reason.
function encodeScript(t){
  return t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function uploadScript(name, text, msg){
  var enc = encodeScript(text);
  var SIZE = 120000;
  var parts = [];
  for (var i = 0; i < enc.length; i += SIZE){ parts.push(enc.substr(i, SIZE)); }
  msg.textContent = "Sending " + parts.length + " parts...";
  post({action: "upload_begin", s_name: name}, function(r0){
    if (!r0 || !r0.success){
      msg.textContent = (r0 && r0.error) || "Could not start the upload.";
      return;
    }
    var i2 = 0;
    function next(){
      if (i2 >= parts.length){
        post({action: "upload_finish", s_name: name}, function(rf){
          if (!rf || !rf.success){
            msg.textContent = "Failed: " + ((rf && rf.error) || "no response");
            return;
          }
          msg.innerHTML = "<b>Wrote " + rf.chars.toLocaleString()
            + " characters to " + esc(rf.name) + ".</b> "
            + "Run Diagnose to confirm.";
        });
        return;
      }
      post({action: "upload_chunk", s_data: parts[i2]}, function(rc){
        if (!rc || !rc.success){
          msg.textContent = "Part " + (i2 + 1) + " failed: "
            + ((rc && rc.error) || "no response")
            + ". Nothing was written.";
          return;
        }
        i2++;
        msg.textContent = "Sent " + i2 + " of " + parts.length + "...";
        next();
      });
    }
    next();
  });
}

function dbDiagnose(){
  var box = $id("stDiag");
  box.innerHTML = "<span class='db-muted'>Looking...</span>";
  post({action: "diagnose_reports"}, function(r){
    if (!r || !r.success){
      box.innerHTML = "<span class='db-err'>" + esc((r && r.error) || "failed")
                    + "</span>";
      return;
    }
    var rows = r.rows || [];
    var exact = [];
    for (var i = 0; i < rows.length; i++){
      if (rows[i].name === r.looking_for) exact.push(rows[i]);
    }
    var h = "<div class='db-muted'>Looking for a script named <b>"
          + esc(r.looking_for) + "</b>.</div>";
    if (!rows.length){
      h += "<div class='db-err'>No stored content matches that name at all. "
        + "The name in Setup does not match what is installed.</div>";
      box.innerHTML = h; return;
    }
    h += "<table class='db-t' style='margin-top:6px;'><thead><tr><th>Id</th>"
      + "<th>Name</th><th>Type</th><th>Archived</th><th>Bytes</th>"
      + "<th>Has list_reports</th><th>Build stamp</th></tr></thead><tbody>";
    for (var j = 0; j < rows.length; j++){
      var w = rows[j];
      var isExact = (w.name === r.looking_for);
      h += "<tr" + (isExact ? " style='font-weight:600;'" : "") + ">"
        + "<td>" + w.id + "</td><td>" + esc(w.name) + "</td>"
        + "<td>" + w.type + "</td><td>" + (w.archived ? "yes" : "no") + "</td>"
        + "<td>" + w.bytes.toLocaleString() + "</td>"
        + "<td>" + (w.has_action ? "yes" : "<span style='color:#a4232b;'>NO</span>") + "</td>"
        + "<td>" + (w.has_build ? "yes" : "<span style='color:#a4232b;'>NO</span>") + "</td>"
        + "</tr>";
    }
    h += "</tbody></table>";

    // TouchPoint takes the FIRST row for the name with no ordering, so more
    // than one is the whole problem.
    if (exact.length > 1){
      h += "<div class='db-err' style='margin-top:8px;'><b>" + exact.length
        + " stored scripts share the name " + esc(r.looking_for) + ".</b> "
        + "TouchPoint runs whichever the database returns first and ignores "
        + "the archived flag, so editing one does not change what executes. "
        + "Rename or delete the ones you do not want.</div>";
    } else if (exact.length === 1 && !exact[0].has_action){
      var e0 = exact[0];
      var known = e0.m1 + e0.m2 + e0.m3;
      h += "<div class='db-err' style='margin-top:8px;'><b>The stored script is "
        + "the old version.</b> It has no list_reports action.<br>";
      if (known === 3){
        // It carries every marker from the version on disk, so
        // it is the same lineage and simply predates the edits.
        h += "It does carry load_person_detail, save_contact_methods and "
          + "eng_person_engagement_scorecard, so it is the same lineage as the "
          + "file on disk and is safe to replace.";
      } else {
        h += "It is missing " + (3 - known) + " of 3 markers from the known "
          + "version, so it may be a different lineage. Save a copy before "
          + "replacing it.";
      }
      h += "</div>";
    } else if (exact.length === 1 && exact[0].has_action && !exact[0].has_build){
      h += "<div class='db-err' style='margin-top:8px;'>Stored copy has "
        + "list_reports but not the newest build stamp, so it is partially "
        + "out of date.</div>";
    } else if (exact.length === 1){
      h += "<div style='margin-top:8px;color:#1b6b33;'><b>The stored script is "
        + "current.</b> If tiles still fail, the problem is elsewhere.</div>";
    } else if (!exact.length){
      // Rows came back, but none is an exact name match: the configured name
      // is close to something installed rather than equal to it.
      var names = [];
      for (var q3 = 0; q3 < rows.length; q3++){ names.push(rows[q3].name); }
      h += "<div class='db-err' style='margin-top:8px;'><b>Nothing is named "
        + "exactly " + esc(r.looking_for) + ".</b> Similar names found: "
        + esc(names.join(", ")) + ". Set the right one in the box above.</div>";
    }
    box.innerHTML = h;
  });
}

var MTYPES = null;

function apprChecked(){
  var out = [], els = document.querySelectorAll(".stAppr");
  for (var i = 0; i < els.length; i++){
    if (els[i].checked) out.push(els[i].value);
  }
  return out.join(",");
}

function dbSettingsSave(){
  var fsel = $id("stFiscal");
  var v = (($id("stName") || {}).value || "").trim();
  post({action: "save_settings", reports_script: v,
        fiscal_month: fsel ? fsel.value : "",
        attendance_scope: (($id("stAttScope") || {}).value || "waag"),
        bg_approval_ok: apprChecked(),
        bg_check_days: (($id("stBgDays") || {}).value || "").trim()}, function(r){
    if (!r || !r.success){ alert((r && r.error) || "Could not save."); return; }
    // What counts as attendance changes every cached attendance number.
    cacheClear();
    $id("stMsg").textContent = "Saved. Reload the page to pick it up.";
  });
}

// How many installed template copies are behind. Shown on the Library button
// so a person finds out without having to go looking.
// Two different things are worth a badge: dashboard templates that have
// moved on, and installed reports whose published definition has changed.
// Both live behind the Library button, so both are counted, and the tooltip
// keeps them apart.
function libUpdateBadge(){
  var behind = 0, staleReports = 0, done = 0;
  function show(){
    if (++done < 2) return;
    var b = $id("libBadge");
    if (!b) return;
    var total = behind + staleReports;
    if (!total){ b.style.display = "none"; return; }
    var bits = [];
    if (behind) bits.push(behind + " dashboard(s) behind their template");
    if (staleReports) bits.push(staleReports + " installed report(s) updated");
    b.textContent = String(total);
    b.style.display = "";
    b.title = bits.join(", ");
  }
  post({action: "list_library"}, function(r){
    if (r && r.success){
      for (var i = 0; i < (r.library || []).length; i++){
        var got = r.library[i].installed || [];
        for (var j = 0; j < got.length; j++){
          var st2 = got[j].state;
          if (st2 === "behind" || st2 === "behind_edited" || st2 === "differs") behind++;
        }
      }
    }
    show();
  });
  post({action: "catalog_browse"}, function(r){
    if (r && r.success){
      var reps = r.reports || [];
      for (var k = 0; k < reps.length; k++){
        if (reps[k].outdated) staleReports++;
      }
    }
    show();
  });
}

function dbLibrary(){
  post({action: "list_library"}, function(r){
    if (!r.success){ alert(r.error); return; }
    // Counts on the tabs, and a red mark on the side that needs attention,
    // so the state is visible without opening each one in turn.
    var nTpl = (r.library || []).length, nBehind = 0;
    for (var q = 0; q < nTpl; q++){
      var gg = r.library[q].installed || [];
      for (var q2 = 0; q2 < gg.length; q2++){
        var s2 = gg[q2].state;
        if (s2 === "behind" || s2 === "behind_edited" || s2 === "differs") nBehind++;
      }
    }
    var dot = "<span style='color:#c0392b;font-weight:700;'> &bull;</span>";
    var h = "<div class='db-tabs' style='margin-bottom:10px;'>"
          + "<button id='libTabBuilt' class='db-tab on' onclick='libSource(0)'>"
          + "Dashboards (" + nTpl + ")" + (nBehind ? dot : "") + "</button>"
          + "<button id='libTabCat' class='db-tab' onclick='libSource(1)'>"
          + "Reports<span id='libCatCount'></span></button></div>"
          + "<p class='db-muted' style='margin:0 0 10px;'>Both come from "
          + "<b>DisplayCache</b>. A <b>dashboard</b> is a ready-made board and "
          + "arrives complete. A <b>report</b> is one set of numbers: "
          + "installing it makes it available, and you put it somewhere "
          + "yourself.</p>"
          + "<div id='libCat' style='display:none;'></div>"
          + "<div id='libBuilt'>"
          + "<p class='db-muted'>Adding a template copies it into your own "
          + "dashboards, so you can change it freely. The template stays as it is.</p>"
          + "<div class='db-grid-cards'>";
    for (var i = 0; i < r.library.length; i++){
      var l = r.library[i];
      var got = l.installed || [];
      h += "<div class='db-card'" + (l.usable ? "" : " style='opacity:.6;'")
        + "><h4>" + esc(l.name)
        + (got.length ? (" <span style='font-size:11px;font-weight:normal;"
                         + "color:#2f6fb5;'>installed</span>") : "")
        + "</h4>"
        + (l.audience ? ("<div class='db-muted'><i>" + esc(l.audience)
                        + "</i></div>") : "")
        + "<div class='db-muted'>" + esc(l.description) + "</div>"
        + "<div class='db-muted'>" + l.tabs + " tab(s), " + l.tiles + " tiles"
        + (l.needs && l.needs.length ? (" &middot; needs " + esc(l.needs.join(", "))) : "")
        + "</div>";
      if (got.length){
        // Named, because adding a template twice makes two dashboards with the
        // same name and nothing on the card to tell them apart.
        h += "<div class='db-muted' style='font-size:12px;margin-top:4px;'>Yours: ";
        for (var g = 0; g < got.length; g++){
          var st = got[g].state;
          h += (g ? ", " : "") + "<a href='#' data-d='" + esc(got[g].id)
            + "' data-s='" + esc(got[g].scope) + "' onclick='libOpen(this);"
            + "return false;'>" + esc(got[g].name) + "</a>"
            // Each state gets its own words: "behind" is safe to replace,
            // "behind and edited" is not, and a copy from before versioning
            // cannot be judged either way.
            + (st === "behind"
               ? " <span style='color:#c0392b;'>(update available)</span>" : "")
            + (st === "behind_edited"
               ? " <span style='color:#c0392b;'>(update available, but you have"
                 + " changed this copy)</span>" : "")
            + (st === "differs"
               ? " <span style='color:#c0392b;'>(differs from the template, "
                 + "installed before updates were tracked so we cannot tell "
                 + "whether that is your change or ours)</span>" : "");
        }
        h += "</div>";
      }
      if (l.usable){
        var anyBehind = false, anyEdited = false;
        for (var b2 = 0; b2 < got.length; b2++){
          if (got[b2].state === "behind") anyBehind = true;
          if (got[b2].state === "behind_edited"){ anyBehind = true; anyEdited = true; }
          // Treated as possibly-edited: we cannot prove it is not, and the
          // costly mistake is overwriting someone's work silently.
          if (got[b2].state === "differs"){ anyBehind = true; anyEdited = true; }
        }
        h += "<div style='margin-top:6px;'>"
          + "<button class='btn btn-xs " + (anyBehind ? "btn-default" : "btn-primary")
          + "' data-l='" + esc(l.id) + "' onclick='libAdd(this)'>"
          + (got.length ? "Add another copy" : "Add") + "</button>"
          + (anyBehind
             ? (" <button class='btn btn-xs btn-primary' data-l='" + esc(l.id)
                + "' data-edited='" + (anyEdited ? "1" : "")
                + "' onclick='libUpdate(this)'>Update mine</button>")
             : "")
          + (got.length
             ? (" <button class='btn btn-xs btn-default' data-l='" + esc(l.id)
                + "' onclick='libReinstall(this)'>Reinstall</button>") : "")
          + "</div>";
        if (anyEdited){
          h += "<div class='db-muted' style='font-size:11px;margin-top:4px;'>"
            + "Updating replaces your copy. Your changes are not merged, so if "
            + "you want to keep them, add a fresh copy alongside and move what "
            + "you need across.</div>";
        }
      } else {
        h += "<div class='db-err' style='font-size:12px;'>" + esc(l.why) + "</div>";
      }
      h += "</div>";
    }
    h += "</div></div>";
    dbModal("DisplayCache Library", h);
    // Fetched now, not on first click: the tab label carries a count and an
    // out-of-date mark, and neither can be shown by a pane nobody opened yet.
    loadCatalogBrowse();
  });
}

function libSource(which){
  var a = $id("libTabBuilt"), b = $id("libTabCat");
  if (a) a.className = "db-tab" + (which === 0 ? " on" : "");
  if (b) b.className = "db-tab" + (which === 1 ? " on" : "");
  $id("libBuilt").style.display = which === 0 ? "" : "none";
  $id("libCat").style.display = which === 1 ? "" : "none";
  // Already loaded when the modal opened; only refetch if that failed.
  if (which === 1 && !CATALOG) loadCatalogBrowse();
}

function loadCatalogBrowse(){
  var box = $id("libCat");
  box.innerHTML = "<span class='db-muted'>Fetching the published catalog...</span>";
  post({action: "catalog_browse"}, function(r){
    if (!r || !r.success){
      box.innerHTML = "<p class='db-err'>" + esc((r && r.error) || "No response")
        + "</p><p class='db-muted'>The catalog is fetched from "
        + "scripts.displaycache.com. If TouchPoint cannot reach it, this stays "
        + "empty and everything else keeps working.</p>";
      return;
    }
    CATALOG = r;
    drawCatalogBrowse();
  });
}

// Reinstalls every installed report whose published hash has moved on.
function catUpdateOutdated(){
  var ids = [];
  var reps = (CATALOG && CATALOG.reports) || [];
  for (var i = 0; i < reps.length; i++){
    if (reps[i].outdated) ids.push(reps[i].id);
  }
  if (!ids.length) return;
  var msg = $id("catUpdMsg");
  if (msg) msg.textContent = "Updating " + ids.length + "...";
  post({action: "catalog_install", ids: ids.join(",")}, function(r){
    if (!r || !r.success){
      alert((r && r.error) || "Could not update.");
      if (msg) msg.textContent = "";
      return;
    }
    // Their cached results were produced by the OLD definitions.
    cacheClear();
    CATALOG = null;
    loadCatalogBrowse();
    // The Library button keeps its own count, and it was left showing the
    // number we had just finished fixing.
    libUpdateBadge();
    alert("Updated " + ((r.installed || []).length) + " report(s). "
        + "Reopen a dashboard to see the new numbers.");
  });
}

function drawCatalogBrowse(){
  var box = $id("libCat");
  var reps = CATALOG.reports || [], dash = CATALOG.dashboards || [];
  var inst = reps.filter(function(x){ return x.installed; }).length;
  var old = reps.filter(function(x){ return x.outdated; }).length;
  // Fill the tab label now that the counts are known.
  var tabc = $id("libCatCount");
  if (tabc){
    tabc.innerHTML = " (" + reps.length + ")"
      + (old ? "<span style='color:#c0392b;font-weight:700;'> &bull;</span>" : "");
  }
  var h = "<p class='db-muted'>" + reps.length + " reports published. <b>"
        + inst + "</b> installed here"
        + (old ? (", <b>" + old + "</b> with a newer version available") : "")
        + ". <b>Add to a dashboard</b> installs a report and takes you "
        + "straight to placing it. <b>Copy only</b> just brings the definition "
        + "across, which changes nothing you can see until you add a tile."
        + "</p>";
  // Installing takes a COPY, so republishing upstream does not reach a report
  // already installed here. The count alone left no way to act on it.
  if (old){
    h += "<div style='margin-bottom:10px;'>"
      + "<button class='btn btn-sm btn-primary' onclick='catUpdateOutdated()'>"
      + "Update " + old + " installed report" + (old === 1 ? "" : "s") + "</button>"
      + " <span id='catUpdMsg' class='db-muted'></span></div>";
  }

  // The published dashboards are the same boards the first tab already
  // offers, through an install path that also resolves widgets by name. Two
  // routes to one thing was the confusion, so this pane is reports only.
  h += "<input id='catFind' placeholder='Filter reports...' "
    + "style='width:100%;padding:6px 10px;margin-bottom:8px;' "
    + "oninput='catFilter(this.value)'>"
    + "<div style='margin-bottom:8px;'>"
    + "<button class='btn btn-xs btn-default' onclick='catInstallVisible()'>"
    + "Install all shown</button></div>"
    + "<div id='catList' style='max-height:320px;overflow:auto;'>";
  for (var j = 0; j < reps.length; j++){
    var rp = reps[j];
    h += "<div class='catRow' data-name='"
      + esc((rp.id + " " + rp.name + " " + rp.category).toLowerCase())
      + "' style='padding:5px 8px;border-bottom:1px solid #f2f5f8;display:flex;"
      + "gap:8px;align-items:center;'>"
      + "<div style='flex:1;'><b>" + esc(rp.name) + "</b>"
      + (rp.installed ? (rp.outdated
          ? " <span class='db-muted'>(update available)</span>"
          : " <span style='color:#1b6b33;font-size:11px;'>installed</span>") : "")
      + "<div class='db-muted'>" + esc(rp.category) + " &middot; "
      + esc(rp.description || "") + "</div></div>"
      + "<button class='btn btn-xs btn-primary' data-r='" + esc(rp.id)
      + "' data-n='" + esc(rp.name) + "' onclick='catAddToDash(this)'>"
      + "Add to a dashboard</button>"
      + " <button class='btn btn-xs btn-default' data-r='" + esc(rp.id)
      + "' onclick='catInstallOne(this)' title='Copy the definition into "
      + "TouchPoint without placing it anywhere. Useful for stocking the list "
      + "before you build; it changes nothing you can see.'>"
      + (rp.installed ? "Refresh copy" : "Copy only") + "</button></div>";
  }
  h += "</div>";
  box.innerHTML = h;
}

function catFilter(v){
  v = (v || "").toLowerCase();
  var rows = document.querySelectorAll("#catList .catRow");
  for (var i = 0; i < rows.length; i++){
    var nm = rows[i].getAttribute("data-name") || "";
    rows[i].style.display = (!v || nm.indexOf(v) >= 0) ? "" : "none";
  }
}

function catInstall(ids, dashId, after, quiet){
  var p = {action: "catalog_install", ids: ids.join(",")};
  if (dashId) p.dash_id = dashId;
  post(p, function(r){
    if (!r || !r.success){
      alert((r && r.error) || "Install failed.");
      return;
    }
    var msg = "Copied " + r.installed.length + " report definition(s) into "
            + "TouchPoint. Nothing has changed on any dashboard: use <b>Add to "
            + "a dashboard</b>, or Edit then Add tile, to place one.";
    if (r.missing && r.missing.length){
      // Anything refused is named. A silent partial install would leave tiles
      // that fail later with no explanation.
      msg += "\\n\\nSkipped:\\n" + r.missing.join("\\n");
    }
    if (!quiet) alert(msg);
    if (after) after(r);
  });
}

// Install a report AND place it, rather than installing it and leaving the
// person to work out that nothing happened. Offers the dashboard they are
// already looking at first, because that is nearly always the answer.
// Install the report, then hand the person to the SAME configure step that
// Add tile uses, on the dashboard they picked, already in edit mode. The
// earlier version added the tile from a little form of its own, which skipped
// filters, aggregation and the settings banner: three things they would then
// have had to go and set anyway.
function catAddToDash(el){
  var rid = el.getAttribute("data-r"), rname = el.getAttribute("data-n") || rid;
  catInstall([rid], "", function(){
    CATALOG = null;
    loadPicker(DASH ? DASH.id : "", true);
    setTimeout(function(){ catChooseDash(rid, rname); }, 400);
  }, true);
}

function catChooseDash(rid, rname){
  var list = DASH_LIST || [];
  if (!list.length){
    alert("Installed " + rname + ", but you have no dashboard to put it on "
        + "yet. Make one with New, then add it there.");
    return;
  }
  var opts = "";
  for (var i = 0; i < list.length; i++){
    opts += "<option value='" + esc(list[i].id) + "'"
         + (DASH && String(list[i].id) === String(DASH.id) ? " selected" : "")
         + ">" + esc(list[i].name) + "</option>";
  }
  dbModal("Add " + rname,
    "<p>Put <b>" + esc(rname) + "</b> on which dashboard?</p>"
    + "<select id='catDashPick' style='width:100%;padding:6px;'>" + opts
    + "</select>"
    + "<div style='margin-top:12px;'>"
    + "<button class='btn btn-sm btn-primary' data-r='" + esc(rid)
    + "' data-n='" + esc(rname) + "' onclick='catPlaceGo(this)'>Continue</button> "
    + "<button class='btn btn-sm btn-default' onclick='dbLibrary()'>Back</button>"
    + "</div>");
}

function catPlaceGo(el){
  var rid = el.getAttribute("data-r"), rname = el.getAttribute("data-n") || rid;
  var did = ($id("catDashPick") || {}).value || "";
  if (!did) return;
  var openConfigure = function(){
    if (!EDIT) dbToggleEdit();
    // dbAddTile loads the catalogs and then shows the picker; dbPickOpen needs
    // those loaded, so it waits for the same list rather than racing it.
    dbAddTile();
    var tries = 0;
    var wait = setInterval(function(){
      tries++;
      var ready = false;
      for (var i = 0; i < (REPORTS || []).length; i++){
        if (REPORTS[i].id === rid) ready = true;
      }
      if (ready){
        clearInterval(wait);
        dbPickOpen(rid, rname);
      } else if (tries > 25){
        clearInterval(wait);
        // The picker is already open; say why it did not jump straight in.
        alert("Installed " + rname + ". Find it in the list to add it.");
      }
    }, 120);
  };
  if (DASH && String(DASH.id) === String(did)){ openConfigure(); return; }
  dbOpen(did, "personal");
  var t2 = 0;
  var waitOpen = setInterval(function(){
    t2++;
    if (DASH && String(DASH.id) === String(did)){
      clearInterval(waitOpen);
      openConfigure();
    } else if (t2 > 25){
      clearInterval(waitOpen);
      alert("Could not open that dashboard.");
    }
  }, 120);
}

function catInstallOne(el){
  catInstall([el.getAttribute("data-r")], "", function(){ loadCatalogBrowse(); });
}

function catInstallVisible(){
  var ids = [], rows = document.querySelectorAll("#catList .catRow");
  for (var i = 0; i < rows.length; i++){
    if (rows[i].style.display === "none") continue;
    var b = rows[i].querySelector("button");
    if (b) ids.push(b.getAttribute("data-r"));
  }
  if (!ids.length){ alert("Nothing shown to install."); return; }
  if (!confirm("Install " + ids.length + " reports?")) return;
  catInstall(ids, "", function(){ loadCatalogBrowse(); });
}

function catInstallDash(el){
  catInstall([], el.getAttribute("data-d"), function(r){
    dbCloseModal();
    if (r.dashboard_id){ DASH_LIST = null; dbOpen(r.dashboard_id, "personal"); }
    else dbHome();
  });
}

function libOpen(el){
  dbCloseModal();
  dbOpen(el.getAttribute("data-d"), el.getAttribute("data-s"));
}

function libAdd(el){ dbAddLib(el.getAttribute("data-l")); }

// Replaces the existing copy rather than adding beside it: the point is to
// pick up template changes, and a second dashboard with the same name is the
// thing people were doing by hand instead.
// Same mechanism as Reinstall, but the wording differs by what is at stake:
// an untouched copy is a routine refresh, an edited one loses work.
function libUpdate(el){
  var id = el.getAttribute("data-l");
  var edited = el.getAttribute("data-edited") === "1";
  var msg = edited
    ? "This template has changed, but you have also changed your copy. "
      + "Updating replaces yours and your changes are lost. Continue?"
    : "Replace your copy with the updated template? You have not changed it, "
      + "so nothing of yours is lost.";
  if (!confirm(msg)) return;
  post({action: "add_from_library", lib_id: id, replace: "1"}, function(r){
    if (!r || !r.success){ alert((r && r.error) || "Could not update."); return; }
    dbCloseModal();
    DASH_LIST = null;
    dbOpen(r.id, "personal");
  });
}

function libReinstall(el){
  var id = el.getAttribute("data-l");
  if (!confirm("Replace your copy with a fresh one from the template? "
             + "Any changes you made to it are lost.")) return;
  post({action: "add_from_library", lib_id: id, replace: "1"}, function(r){
    if (!r || !r.success){ alert((r && r.error) || "Could not reinstall."); return; }
    dbCloseModal();
    DASH_LIST = null;
    dbOpen(r.id, "personal");
  });
}

function dbAddLib(id){
  post({action: "add_from_library", lib_id: id}, function(r){
    if (!r.success){ alert(r.error); return; }
    if (r.missing && r.missing.length){
      // Templates name widgets rather than numbering them, so one this church
      // does not have is skipped. Better said out loud than silently absent.
      alert("Added, but these widgets are not present in your TouchPoint and "
          + "were skipped:\\n\\n" + r.missing.join("\\n"));
    }
    dbCloseModal();
    DASH_LIST = null;
    dbOpen(r.id, "personal");
  });
}

// ---------- open / render ----------
var OPEN_TAB = 0;

function dbOpen(id, scope, tab){
  OPEN_TAB = tab || 0;
  var gp = {action: "get_dashboard", d_id: id, d_scope: scope || ""};
  if (ADMIN_OWNER) gp.owner_uid = ADMIN_OWNER.uid;
  post(gp, function(r){
    if (!r.success){ alert(r.error); return; }
    DASH = r.dashboard;
    DASH.can_edit = r.can_edit;
    TAB = 0;
    if (OPEN_TAB && OPEN_TAB < (DASH.tabs || []).length){ TAB = OPEN_TAB; }
    OPEN_TAB = 0;
    EDIT = false;
    $id("dbHome").style.display = "none";
    $id("dbGrid").style.display = "";
    $id("dbTitle").textContent = DASH.name || "Dashboard";
    var who = r.owner_name || "";
    $id("dbScope").textContent = who
      ? ("(" + who + "'s personal dashboard)")
      : (DASH.scope === "shared" ? "(shared)" : "(personal)");
    $id("dbScope").style.color = who ? "#a8442a" : "";
    $id("dbEditBtn").style.display = r.can_edit ? "" : "none";
    $id("dbEditBtn").textContent = "Edit";
    $id("dbEditTools").style.display = "none";
    var sb2 = $id("dbShareBtn");
    if (sb2) sb2.style.display = "";
    writeUrlState();
    var cs = $id("dbCacheMins");
    if (cs) cs.value = String(cacheMinutes());
    var rb = $id("dbRefreshBtn");
    if (rb) rb.style.display = cacheMinutes() ? "" : "none";
    var rs = $id("dbRange");
    if (rs){ rs.style.display = ""; rs.value = currentRange(); }
    fillScopeSelect();
    fillProgramSelect();
    fillDefQuery();
    applyCapsToBar();
    libUpdateBadge();
    cacheSweep();
    renderTabs();
    renderGrid();
    narrowGuard();
    loadPicker(DASH.id, false);
  });
}

function renderTabs(){
  applyCapsToBar();
  var tabs = DASH.tabs || [];
  var el = $id("dbTabs");
  if (tabs.length <= 1 && !EDIT){ el.style.display = "none"; }
  else { el.style.display = "flex"; }
  var h = "";
  for (var i = 0; i < tabs.length; i++){
    h += "<button class='db-tab" + (i === TAB ? " on" : "") + "' onclick='dbTab("
      + i + ")'>" + esc(tabs[i].name || ("Tab " + (i + 1))) + "</button>";
  }
  el.innerHTML = h;
}

function dbTab(i){
  TAB = i;
  writeUrlState();
  renderTabs();
  renderGrid();
}

function tiles(){
  var t = (DASH.tabs || [])[TAB];
  return (t && t.tiles) || [];
}

var RENDERING = false;

function renderGrid(){
  if (RENDERING){
    // Re-entered from narrowGuard part-way through a render. Let the pass in
    // progress finish; it will apply the current EDIT and column state anyway.
    return;
  }
  RENDERING = true;
  try { renderGridInner(); } finally { RENDERING = false; }
}

function renderGridInner(){
  for (var k in CHARTS){
    try { CHARTS[k].destroy(); } catch (e) { }
  }
  CHARTS = {};
  for (var o = 0; o < OBSERVERS.length; o++){
    try { OBSERVERS[o].disconnect(); } catch (e) { }
  }
  OBSERVERS = [];
  AUTHORED_H = {};
  LAST_PAYLOAD = {};
  SEARCH_IDS = {};
  var host = $id("dbGrid");
  if (GRID){ try { GRID.destroy(false); } catch (e) { } GRID = null; }
  host.innerHTML = "";

  var list = tiles();
  if (!list.length){
    host.innerHTML = "<p class='db-muted'>No tiles on this tab."
      + (DASH.can_edit ? " Click <b>Edit</b> then <b>Add tile</b>." : "") + "</p>";
    return;
  }

  for (var i = 0; i < list.length; i++){
    var t = list[i];
    var item = document.createElement("div");
    item.className = "grid-stack-item";
    item.setAttribute("gs-x", t.x || 0);
    item.setAttribute("gs-y", t.y || 0);
    item.setAttribute("gs-w", t.w || 4);
    item.setAttribute("gs-h", t.h || 4);
    item.setAttribute("data-idx", i);
    // gs-id is what GRID.save() reports back, and the harvest fallback maps
    // by it. Without it that fallback silently persists nothing.
    item.setAttribute("gs-id", i);
    item.innerHTML =
      "<div class='grid-stack-item-content'><div class='db-tile'>"
      + "<div class='db-tile-hd'><span>" + esc(t.title || t.report_id) + "</span>"
      + "<span class='db-age' id='age" + i + "'></span>"
      + "<span class='db-age' id='scope" + i + "'></span>"
      + pinnedHtml(t)
      + "<button class='db-tile-exp' onclick='dbExpand(" + i + ")' "
      + "title='Open this tile larger, with the data behind it'>&#9974;</button>"
      + (EDIT ? "<button class='db-tile-fit" + (t.nocache ? "" : " on")
                + "' style='margin-left:4px;"
                + "' onclick='dbToggleCache(" + i + ")' title='"
                + (t.nocache ? "Always fetched fresh. Click to allow caching."
                             : "Uses the dashboard cache. Click to always fetch fresh.")
                + "' style='margin-left:6px;'>"
                + (t.nocache ? "live" : "cached") + "</button>" : "")
      + (EDIT ? "<button class='db-tile-fit" + (fitOn(t) ? " on" : "")
                + "' style='margin-left:4px;"
                + "' onclick='dbToggleFit(" + i + ")' title='"
                + (fitOn(t) ? "Grows to fit its content. Click for a fixed height."
                            : "Fixed height, content scrolls. Click to grow to fit.")
                + "'>" + (fitOn(t) ? "auto" : "fixed") + "</button>"
                + "<button class='db-tile-fit' style='margin-left:4px;' "
                + "onclick='dbEditTile(" + i + ")' "
                + "title='Rename, filter, or change how this tile shows'>"
                + "edit</button>"
                + "<button class='db-tile-x' onclick='dbRemoveTile(" + i
                + ")' title='Remove tile'>&times;</button>" : "")
      + "</div><div class='db-tile-bd' id='tb" + i + "'>"
      + "<span class='db-muted'>Loading...</span></div></div></div>";
    host.appendChild(item);
  }

  if (typeof GridStack === "undefined"){
    // Church networks do block CDNs. Say so, and lay the tiles out plainly
    // rather than showing an empty page with nothing to explain it.
    var warn = document.createElement("div");
    warn.className = "db-err";
    warn.style.marginBottom = "8px";
    warn.innerHTML = "Layout library did not load (CDN blocked?). "
                   + "Tiles are shown stacked and cannot be moved.";
    host.insertBefore(warn, host.firstChild);
    host.style.display = "block";
    var kids = host.querySelectorAll(".grid-stack-item");
    for (var g = 0; g < kids.length; g++){
      kids[g].style.position = "static";
      kids[g].style.height = "320px";
      kids[g].style.marginBottom = "10px";
    }
  } else {
    GRID = GridStack.init({
      column: 12, cellHeight: CELL_H, margin: CELL_MARGIN,
      disableDrag: !EDIT, disableResize: !EDIT,
      handle: ".db-tile-hd"
    }, host);
    GRID_COLS = 12;
    applyColumns();
    GRID.on("resizestart", function(){ gridGrabbing(true); });
    GRID.on("dragstart", function(){ gridGrabbing(true); });
    GRID.on("resizestop", function(ev, el){
      gridGrabbing(false);
      var idx = parseInt(el.getAttribute("data-idx"), 10);
      if (!isNaN(idx) && el.gridstackNode){
        // Recorded BEFORE fit runs again, so auto-grow treats the height the
        // user just chose as the floor rather than undoing it.
        AUTHORED_H[idx] = el.gridstackNode.h;
        setTimeout(function(){ autoGrowTile(idx); }, 60);
      }
      if (EDIT) markDirty();
      harvestLayout();
    });
    GRID.on("dragstop", function(){
      gridGrabbing(false);
      if (EDIT) markDirty();
      harvestLayout();
    });
  }

  // Tiles load one at a time. Firing a dozen report queries at once is how a
  // dashboard takes the database down; this keeps it to one in flight.
  loadNext(0);
}

function loadNext(i){
  var list = tiles();
  if (i >= list.length) return;
  loadTile(i, function(){ loadNext(i + 1); });
}

function loadTile(i, done){
  var t = tiles()[i];
  var box = $id("tb" + i);
  if (!box){ if (done) done(); return; }
  var hit = cacheGet(t);
  if (hit){
    renderFromPayload(t, box, i, hit.v);
    ageLabel(i, hit.at);
    scopeLabel(i, hit.v);
    needsSettingBar(box, hit.v);
    if (done) done();
    return;
  }
  ageLabel(i, 0);
  if (t.kind === "links"){ renderLinksTile(t, box); if (done) done(); return; }
  if (t.kind === "opscheck"){ loadOpsTile(t, box, done, i); return; }
  if (t.kind === "widget"){ loadWidgetTile(t, box, done, i); return; }
  if (t.kind === "custom"){ loadCustomTile(t, box, done, i); return; }
  var params = {action: "run_report", report_id: t.report_id,
                display_type: t.display || "chart"};
  for (var k in (t.filters || {})){ params[k] = t.filters[k]; }
  var gr0 = currentRange();
  if (gr0 && !t.own_dates && !(t.filters || {}).date_range){
    params.date_range = gr0;
  }
  if (t.serving_types){ params.serving_types = t.serving_types; }
  if (t.query){
    withSearchIds(t.query, function(ids, err){
      if (err){
        box.innerHTML = "<span class='db-err'>" + esc(err) + "</span>";
        if (done) done();
        return;
      }
      params.people_ids = ids.join(",");
      runReportTile(t, box, i, params, done);
    });
    return;
  }
  runReportTile(t, box, i, params, done);
}

// Saved searches are resolved once per page, not once per tile: several tiles
// commonly share one, and each resolution is a real query.
var SEARCH_IDS = {};

function withSearchIds(name, cb){
  if (SEARCH_IDS[name]){ cb(SEARCH_IDS[name], ""); return; }
  post({action: "resolve_search", q_name: name}, function(r){
    if (!r || !r.success){
      cb([], (r && r.error) || "Could not run the search.");
      return;
    }
    SEARCH_IDS[name] = r.ids || [];
    cb(SEARCH_IDS[name], "");
  });
}

function runReportTile(t, box, i, params, done){
  if (tileSource(t) === "catalog"){ runCatalogTile(t, box, i, params, done); return; }
  postReports(params, function(r){
    if (!r || !r.success){
      box.innerHTML = "<span class='db-err'>"
        + esc((r && r.error) || "No response") + "</span>";
      if (done) done();
      return;
    }
    // Only successful payloads are cached. Caching an error would keep showing
    // it for the whole TTL even after the cause was fixed.
    cachePut(t, r);
    LAST_PAYLOAD[i] = r;
    renderReport(t, box, i, r);
    if (done) done();
  });
}

// A catalog tile is executed by THIS script against the locally installed
// definition, so it works with no other script present.
function runCatalogTile(t, box, i, params, done){
  var p = {action: "run_catalog", report_id: t.report_id};
  var gr = currentRange();
  // A tile that pinned its own dates keeps them; the global range only fills
  // the gap, and the server ignores it for reports with no date parameter.
  if (gr && !t.own_dates){ p.global_range = gr; }
  var gq = currentQuery();
  if (gq && !t.own_scope){ p.global_query = gq; }
  var gpr = currentProgram();
  if (gpr && !t.own_area){ p.global_program = gpr; }
  if (t.query){ p.q_name = t.query; }
  if (t.serving_types){ p.serving_types = t.serving_types; }
  if (t.filters && JSON.stringify(t.filters) !== "{}"){
    p.filters = encPayload(JSON.stringify(t.filters));
  }
  post(p, function(r){
    if (!r || !r.success){
      box.innerHTML = (r && r.restricted)
        ? restrictedHtml()
        : ("<span class='db-err'>" + esc((r && r.error) || "No response")
           + "</span>");
      if (done) done();
      return;
    }
    cachePut(t, r);
    LAST_PAYLOAD[i] = r;
    scopeLabel(i, r);
    renderCatalogResult(t, box, i, r);
    needsSettingBar(box, r);
    if (done) done();
  });
}

// Grouping and measuring happen here rather than in SQL: the report is a
// fixed definition shared across churches, so reshaping it for a particular
// tile is the tile's business.
function aggName(a){
  if (a.fn === "count") return "Count";
  if (a.fn === "none") return a.value || "Value";
  return a.fn.charAt(0).toUpperCase() + a.fn.slice(1) + " of " + a.value;
}

function numOf(v){
  if (typeof v === "number") return v;
  // Strips currency and thousands separators, which arrive as formatted text.
  var n = parseFloat(String(v === null || v === undefined ? "" : v)
                     .replace(/[$,%\s]/g, ""));
  return isNaN(n) ? 0 : n;
}

function fmtNum(n){
  if (n === null || n === undefined || isNaN(n)) return "0";
  // Averages want a decimal; counts and sums read better without one.
  var r = (Math.abs(n - Math.round(n)) > 0.001) ? n.toFixed(1) : Math.round(n);
  return Number(r).toLocaleString();
}

function reduceRows(rows, fn, col){
  if (fn === "count") return rows.length;
  var vals = [];
  for (var i = 0; i < rows.length; i++){ vals.push(numOf(rows[i][col])); }
  if (!vals.length) return 0;
  if (fn === "sum" || fn === "none"){
    var s = 0;
    for (var a = 0; a < vals.length; a++){ s += vals[a]; }
    return s;
  }
  if (fn === "avg"){
    var t = 0;
    for (var b = 0; b < vals.length; b++){ t += vals[b]; }
    return t / vals.length;
  }
  if (fn === "min"){ return Math.min.apply(null, vals); }
  if (fn === "max"){ return Math.max.apply(null, vals); }
  return 0;
}

function aggregate(rows, a){
  var keys = [], buckets = {};
  for (var i = 0; i < rows.length; i++){
    var k = rows[i][a.label];
    k = (k === null || k === undefined || k === "") ? "(blank)" : String(k);
    if (!buckets[k]){ buckets[k] = []; keys.push(k); }
    buckets[k].push(rows[i]);
  }
  var out = [];
  for (var j = 0; j < keys.length; j++){
    out.push({k: keys[j], v: reduceRows(buckets[keys[j]], a.fn, a.value)});
  }
  // "none" means the query already aggregated, so its own row order is
  // meaningful (day of week, age band) and must not be re-sorted.
  if (a.fn !== "none"){
    out.sort(function(x, y){ return y.v - x.v; });
  }
  var top = (a.top === "all") ? out.length : (parseInt(a.top, 10) || 15);
  if (out.length > top){
    var rest = 0;
    for (var m = top; m < out.length; m++){ rest += out[m].v; }
    out = out.slice(0, top);
    out.push({k: "Other (" + (keys.length - top) + ")", v: rest});
  }
  var labels = [], values = [];
  for (var n = 0; n < out.length; n++){ labels.push(out[n].k); values.push(out[n].v); }
  return {labels: labels, values: values};
}

// Table state per tile, so a sort survives a redraw without re-running the
// report. Keyed by tile index, same as LAST_PAYLOAD.
var TABLE_STATE = {};

// A person column means the rows are people, and people can be acted on. The
// id column itself is not shown: it is the checkbox.
var PID_COLS = ["PeopleId", "PeopleID", "peopleid"];

function peopleCol(cols){
  for (var i = 0; i < cols.length; i++){
    for (var j = 0; j < PID_COLS.length; j++){
      if (cols[i] === PID_COLS[j]) return cols[i];
    }
  }
  return "";
}

// Numbers sort as numbers and dates as dates; everything else falls back to a
// case-insensitive string compare. Sorting a column of "1,234" as text would
// put it between 1 and 2.
function cellSortValue(v){
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return v;
  var str = String(v);
  if (/^-?[\d,]+(\.\d+)?$/.test(str.replace(/[$\s]/g, ""))){
    var n = parseFloat(str.replace(/[$,\s]/g, ""));
    if (!isNaN(n)) return n;
  }
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)){
    var p = str.split("/");
    return new Date(+p[2], +p[0] - 1, +p[1]).getTime();
  }
  return str.toLowerCase();
}

function sortRows(rows, col, dir){
  var out = rows.slice();
  out.sort(function(a, b){
    var x = cellSortValue(a[col]), y = cellSortValue(b[col]);
    // Blanks sort last either way, so a sort never buries the real values.
    if (x === null && y === null) return 0;
    if (x === null) return 1;
    if (y === null) return -1;
    if (typeof x === "number" && typeof y === "number"){ return (x - y) * dir; }
    x = String(x); y = String(y);
    return (x < y ? -1 : (x > y ? 1 : 0)) * dir;
  });
  return out;
}

// Raises this table's own row cap. Kept per tile rather than global so one
// long list does not slow every other tile on the board.
// The key arrives as text from the attribute, but TABLE_STATE is keyed by
// number for tiles. Coerced back so both forms find their state.
function catShowMoreEl(el){
  var k = el.getAttribute("data-key");
  catShowMore(/^[0-9]+$/.test(k) ? parseInt(k, 10) : k);
}

function catShowMore(i){
  var st = TABLE_STATE[i];
  if (!st) return;
  st.limit = (st.limit || 200) + 500;
  drawCatalogTable(i);
}

function catalogSortEl(el){
  var k = el.getAttribute("data-key");
  catalogSort(/^[0-9]+$/.test(k) ? parseInt(k, 10) : k,
              parseInt(el.getAttribute("data-col"), 10));
}

function catalogSort(i, colIdx){
  var st = TABLE_STATE[i];
  if (!st) return;
  // Rebuild the visible column list the same way the draw does, so the index
  // means the same thing on both sides.
  var pid = peopleCol(st.cols), shown = [];
  for (var a = 0; a < st.cols.length; a++){
    if (st.cols[a] !== pid) shown.push(st.cols[a]);
  }
  var col = shown[colIdx];
  if (!col) return;
  // Same column toggles direction; a new column starts ascending.
  if (st.sort && st.sort.col === col){
    st.sort = {col: col, dir: -st.sort.dir};
  } else {
    st.sort = {col: col, dir: 1};
  }
  drawCatalogTable(i);
}

function drawCatalogTable(i){
  var st = TABLE_STATE[i];
  if (!st || !st.box) return;
  var cols = st.cols, rows = st.rows;
  if (st.sort){ rows = sortRows(rows, st.sort.col, st.sort.dir); }
  var pid = peopleCol(cols);
  var shown = [];
  for (var a = 0; a < cols.length; a++){
    if (cols[a] !== pid) shown.push(cols[a]);
  }

  // Computed before the header, which names it in the select-all tooltip.
  var cap = Math.min(rows.length, st.limit || 200);
  st.selected = st.selected || {};
  var h = "";
  if (pid){
    h += "<div class='db-people' data-sel='"
      + esc(String(i)) + "'>" + peopleActionBar();
  }
  h += "<div class='db-tw'><table class='db-t'><thead><tr>";
  if (pid){
    h += "<th style='width:26px;'><input type='checkbox' class='pkAll' "
      + "title='Ticks all " + rows.length + " loaded rows, not only those shown' "
      + "onclick='pkToggleAll(this)'></th>";
  }
  for (var c = 0; c < shown.length; c++){
    var arrow = "";
    if (st.sort && st.sort.col === shown[c]){
      arrow = st.sort.dir > 0 ? " \u25b2" : " \u25bc";
    }
    // Column INDEX not name, and the table key as an attribute rather than
    // inlined: the key is a number for a tile but the string "exp" in the
    // expanded view, and inlining that produced catalogSort(exp,0), a bare
    // identifier that threw on every click.
    h += "<th style='cursor:pointer;white-space:nowrap;' data-key='" + esc(i)
      + "' data-col='" + c + "' onclick='catalogSortEl(this)'>"
      + esc(shown[c]) + arrow + "</th>";
  }
  h += "</tr></thead><tbody>";

  for (var r = 0; r < cap; r++){
    h += "<tr>";
    if (pid){
      var rid = String(rows[r][pid]);
      h += "<td><input type='checkbox' class='pkOne' value='" + esc(rid) + "'"
        + (st.selected[rid] ? " checked" : "")
        + " onclick='pkPick(this)'></td>";
    }
    for (var c2 = 0; c2 < shown.length; c2++){
      var v = rows[r][shown[c2]];
      // The name links to the person, since the id column is now a checkbox.
      if (pid && (shown[c2] === "Name" || shown[c2] === "Name2")){
        h += "<td><a href='" + esc(personHref(rows[r][pid], LAST_PAYLOAD[i]))
          + "' target='_blank'>" + esc(v) + "</a></td>";
      } else {
        h += "<td>" + esc(v) + "</td>";
      }
    }
    h += "</tr>";
  }
  h += "</tbody></table></div>";
  // Say plainly what is on screen, what was loaded, and whether the server
  // stopped early. "Select all" can only ever tick the rows that are drawn,
  // and it previously did that silently on a 200-row slice of a longer list.
  var pay = LAST_PAYLOAD[i] || {};
  if (rows.length > cap || pay.truncated){
    h += "<div class='db-muted' style='margin-top:6px;'>Showing <b>" + cap
      + "</b> of <b>" + rows.length + "</b> loaded"
      + (pay.truncated
         ? ", and the report returned more than that so the rest were not read"
         : "")
      + ". Sorting and <b>select all</b> cover everything loaded"
      + (pid ? "; individual ticks apply to the rows shown" : "") + ".";
    if (rows.length > cap){
      h += " <button class='btn btn-xs btn-default' data-key='" + esc(i)
        + "' onclick='catShowMoreEl(this)'>Show more</button>";
    }
    h += "</div>";
  }
  if (pid){ h += "</div>"; }
  st.box.firstChild.innerHTML = h;
  if (pid){
    var bar = st.box.firstChild.querySelector(".pkCountLbl");
    if (bar) pkCount(bar);
  }
}

// Catalog results are plain columns and rows, so they are shaped into whatever
// the tile is set to display rather than arriving pre-rendered.
function renderCatalogResult(t, box, i, r){
  var cols = r.columns || [], rows = r.rows || [];
  var disp = t.display || (r.display && r.display.default) || "table";
  if (!rows.length){
    // A count of nothing is zero, not missing data. Exception reports are
    // empty precisely when everything is fine, and "No rows." reads as though
    // the tile is broken on exactly the day it should read as reassurance.
    if (disp === "kpi" && t.agg
        && (t.agg.fn === "count" || t.agg.fn === "sum")){
      box.innerHTML = "<div class='db-fit'><div class='db-kpi'>0</div>"
        + "<div class='db-muted'>" + esc(t.agg.label || "Count") + "</div></div>";
      return;
    }
    box.innerHTML = "<span class='db-muted'>None right now.</span>";
    return;
  }
  if (disp === "chart"){
    var d = r.display || {};
    var a = t.agg;
    if (a && a.label){
      var g = aggregate(rows, a);
      drawChart(box, i, {type: a.chart || d.chart_type || "bar",
                         labels: g.labels,
                         datasets: [{label: aggName(a), data: g.values}]});
      return;
    }
    // No tile-level setup and no chart configuration in the definition: this
    // is a detail report, and charting its first column plots record ids
    // against nothing. Say so instead of drawing an empty grid.
    if (!d.chart_label_col){
      box.innerHTML = "<div class='db-muted' style='padding:12px;'>"
        + "This report has no chart built in, so the tile needs to be told "
        + "what to group by and what to measure. Remove it and add it again "
        + "as a chart to set that up."
        + "</div>";
      return;
    }
    var labelCol = d.chart_label_col;
    var valCols = d.chart_data_cols && d.chart_data_cols.length
                ? d.chart_data_cols
                : [cols.filter(function(c){ return c !== labelCol; })[0]];
    var labels = [], sets = [];
    for (var v = 0; v < valCols.length; v++){ sets.push({label: valCols[v], data: []}); }
    for (var j = 0; j < rows.length; j++){
      labels.push(rows[j][labelCol]);
      for (var k = 0; k < valCols.length; k++){
        sets[k].data.push(Number(rows[j][valCols[k]]) || 0);
      }
    }
    drawChart(box, i, {type: d.chart_type || "bar", labels: labels, datasets: sets});
    return;
  }
  if (disp === "kpi"){
    var ka = t.agg;
    var kval, klabel, more = "";
    if (ka && ka.fn){
      kval = reduceRows(rows, ka.fn, ka.value);
      klabel = aggName(ka);
      // The server stopped reading at its cap, so every aggregate over these
      // rows is a floor rather than a total.
      if (r.truncated){ more = "+"; klabel = "at least this many"; }
    } else {
      var first = rows[0];
      var kcol = cols.filter(function(c){
        return typeof first[c] === "number"; })[0] || cols[0];
      kval = Number(first[kcol]);
      klabel = kcol;
    }
    box.innerHTML = "<div class='db-fit'><div class='db-kpi'>"
      + esc(fmtNum(kval)) + esc(more) + "</div>"
      + "<div class='db-muted'>" + esc(klabel) + "</div></div>";
    return;
  }
  box.innerHTML = "<div class='db-fit'></div>";
  TABLE_STATE[i] = {cols: cols, rows: rows, sort: (TABLE_STATE[i] || {}).sort,
                    box: box, tile: t};
  drawCatalogTable(i);
  if (fitOn(t)){
    watchTile(i);
    setTimeout(function(){ autoGrowTile(i); }, 200);
  }
}

function renderReport(t, box, i, r){
  if (r.display === "chart" && r.chart_data){
    drawChart(box, i, r.chart_data);
    return;
  }
  // The same table the catalog path draws, so a reporting-sourced tile gets
  // sorting and the people actions too instead of a static grid.
  if (r.display !== "kpi" && r.grid_data){
    var g = normalizeGrid(r.grid_data);
    box.innerHTML = "<div class='db-fit'></div>";
    TABLE_STATE[i] = {cols: g.cols, rows: g.rows,
                      sort: (TABLE_STATE[i] || {}).sort, box: box, tile: t};
    drawCatalogTable(i);
    if (fitOn(t)){
      watchTile(i);
      setTimeout(function(){ autoGrowTile(i); }, 200);
    }
    return;
  }
  var inner = "<span class='db-muted'>Nothing to show.</span>";
  if (r.display === "kpi" && r.html){ inner = r.html; }
  box.innerHTML = "<div class='db-fit'></div>";
  box.firstChild.innerHTML = inner;
  if (fitOn(t)){
    watchTile(i);
    setTimeout(function(){ autoGrowTile(i); }, 200);
  }
}

function fmtVal(v, money){
  if (money){
    return "$" + Number(v).toLocaleString(undefined,
      {minimumFractionDigits: 0, maximumFractionDigits: 0});
  }
  var r = Math.round(Number(v) * 10) / 10;
  return Number(r).toLocaleString();
}

function renderCustom(box, key, r, disp){
  var rows = r.rows || [];
  var note = r.note ? ("<div class='db-muted' style='margin-top:6px;'>"
                     + esc(r.note) + "</div>") : "";
  if (!rows.length){
    box.innerHTML = "<span class='db-muted'>Nothing matched.</span>" + note;
    return;
  }
  // A single total has nothing to plot, whatever shape was asked for.
  if (r.single || disp === "number"){
    box.innerHTML = "<div class='db-fit'><div class='db-kpi'>"
      + esc(fmtVal(rows[0].value, r.money)) + "</div>"
      + "<div class='db-muted'>" + esc(r.measure_label || "") + "</div>"
      + note + "</div>";
    return;
  }
  if (disp === "table"){
    var h = "<table class='db-t'><thead><tr><th>"
          + esc(r.dimension_label || "Group") + "</th><th style='text-align:right;'>"
          + esc(r.measure_label || "Value") + "</th></tr></thead><tbody>";
    for (var i = 0; i < rows.length; i++){
      h += "<tr><td>" + esc(rows[i].label) + "</td>"
        + "<td style='text-align:right;'>" + esc(fmtVal(rows[i].value, r.money))
        + "</td></tr>";
    }
    h += "</tbody></table>" + note;
    box.innerHTML = "<div class='db-fit'></div>";
    box.firstChild.innerHTML = h;
    return;
  }
  var labels = [], vals = [];
  for (var j = 0; j < rows.length; j++){
    labels.push(rows[j].label); vals.push(rows[j].value);
  }
  drawChart(box, key, {type: disp, labels: labels,
                       datasets: [{label: r.measure_label || "", data: vals}]});
}

// Links a person collects for one dashboard. Deliberately part of the
// dashboard rather than a TouchPoint widget: a widget has to be deployed and
// role-assigned by an admin and is shared church-wide, while these travel with
// the board and differ for every church that installs it.
// Folded groups are remembered per person, per tile: it is a viewing
// preference, not a property of the dashboard everyone shares.
function linkFoldKey(idx, cat){
  return "tpxidb:fold:" + (DB_BOOT.user_id || 0) + ":"
       + ((DASH && DASH.id) || "") + ":" + idx + ":" + cat;
}

function linkFolded(idx, cat){
  try { return cacheStore().getItem(linkFoldKey(idx, cat)) === "1"; }
  catch (e) { return false; }
}

function linkFold(el){
  var idx = parseInt(el.getAttribute("data-t"), 10);
  var cat = el.getAttribute("data-c");
  try {
    var k = linkFoldKey(idx, cat);
    var st = cacheStore();
    if (st.getItem(k) === "1") st.removeItem(k); else st.setItem(k, "1");
  } catch (e) { }
  var t = tiles()[idx], box = $id("tb" + idx);
  if (t && box) renderLinksTile(t, box);
}

function linkVisible(l){
  var need = l.roles || "";
  if (!need) return true;
  var mine = DB_BOOT.roles || [];
  var want = String(need).split(",");
  for (var i = 0; i < want.length; i++){
    var w = want[i].trim();
    if (!w) continue;
    for (var j = 0; j < mine.length; j++){
      if (mine[j] === w) return true;
    }
  }
  return false;
}

function renderLinksTile(t, box){
  var all = t.links || [];
  // Shown in the order the rows were arranged in the editor.
  var list = [];
  for (var a = 0; a < all.length; a++){
    if (linkVisible(all[a]) && safeUrl(all[a].url)) list.push(all[a]);
  }
  if (!list.length){
    box.innerHTML = "<span class='db-muted'>"
      + (all.length ? "Nothing here is available to you."
                    : "No links yet. Use <b>edit</b> on this tile to add some.")
      + "</span>";
    return;
  }
  // Grouped by category, in the order the categories first appear.
  var order = [], groups = {};
  for (var i = 0; i < list.length; i++){
    var c = String(list[i].cat || "");
    if (!groups[c]){ groups[c] = []; order.push(c); }
    groups[c].push(list[i]);
  }
  var idx = TILE_IDX_OF(box);
  var h = "<div class='db-fit'>";
  for (var g = 0; g < order.length; g++){
    var cat = order[g];
    var shut = linkFolded(idx, cat);
    // A single unnamed group needs no heading; it would just be a blank line.
    if (cat){
      h += "<div class='db-linkcat' data-t='" + idx + "' data-c='" + esc(cat)
        + "' onclick='linkFold(this)' title='Show or hide this group'>"
        + "<i class='fa fa-caret-" + (shut ? "right" : "down") + "'></i> "
        + esc(cat) + " <span class='db-muted'>(" + groups[cat].length
        + ")</span></div>";
      if (shut) continue;
    }
    h += "<div class='db-links'>";
    var arr = groups[cat];
    for (var k = 0; k < arr.length; k++){
      var u = safeUrl(arr[k].url);
      h += "<a class='db-link' href='" + esc(u) + "'"
        + (arr[k].newtab ? " target='_blank' rel='noopener'" : "") + ">"
        + "<div class='db-link-ico'><i class='fa "
        + esc(arr[k].icon || "fa-link") + "'></i></div>"
        + "<div class='db-link-lbl'><span>" + esc(arr[k].label || u)
        + "</span></div></a>";
    }
    h += "</div>";
  }
  h += "</div>";
  box.innerHTML = h;
  if (fitOn(t)){ watchTile(TILE_IDX_OF(box)); }
}

// A dashboard can be shared, and its JSON is edited by people. Anything but a
// plain web address or a path within this site is refused rather than
// rendered: "javascript:" in an href is a script someone else runs as you.
function safeUrl(u){
  u = String(u || "").trim();
  if (!u) return "";
  // A leading // is protocol-relative and leaves this site entirely, while
  // looking like an internal path to anyone reading the dashboard.
  if (u.substring(0, 2) === "//") return "";
  if (u.charAt(0) === "/") return u;              // path on this TouchPoint
  if (/^https?:\/\//i.test(u)) return u;
  // No scheme given: treat it as a host, which is what people type.
  if (/^[\w.-]+\.[a-z]{2,}(\/|$)/i.test(u)) return "https://" + u;
  return "";
}

function TILE_IDX_OF(box){
  var id = box && box.id ? box.id : "";
  return parseInt(id.replace("tb", ""), 10);
}

function loadCustomTile(t, box, done, idx){
  var cp = {action: "run_custom", spec: encPayload(JSON.stringify(t.spec || {}))};
  var gr1 = currentRange();
  if (gr1 && !t.own_dates){ cp.global_range = gr1; }
  var gq1 = currentQuery();
  if (gq1 && !t.own_scope){ cp.global_query = gq1; }
  post(cp, function(r){
    if (!r || !r.success){
      box.innerHTML = (r && r.restricted)
        ? restrictedHtml()
        : ("<span class='db-err'>" + esc((r && r.error) || "No response")
           + "</span>");
      if (done) done();
      return;
    }
    cachePut(t, r);
    LAST_PAYLOAD[idx] = r;
    renderCustom(box, idx, r, (t.spec && t.spec.display) || "bar");
    if (fitOn(t)){
      watchTile(idx);
      setTimeout(function(){ autoGrowTile(idx); }, 200);
    }
    if (done) done();
  });
}

// One place that turns a stored payload back into a rendered tile, so a cache
// hit and a fresh fetch cannot drift apart.
var LAST_PAYLOAD = {};

// EVERY tile kind that caches must be handled here, not only on the fetch
// path. A kind missing from this function does not fail loudly: it falls
// through to renderReport, which finds no grid_data and prints "Nothing to
// show" over a cached result that was fine. That is what happened to check
// tiles, and to catalog tiles before them.
function renderFromPayload(t, box, idx, payload){
  LAST_PAYLOAD[idx] = payload;
  if (t.kind === "links"){ renderLinksTile(t, box); return; }
  if (t.kind === "opscheck"){
    scopeLabel(idx, payload);
    renderOpsResult(t, box, idx, payload);
    return;
  }
  if (t.kind === "widget"){
    box.innerHTML = "<div class='db-fit'></div>";
    box.firstChild.innerHTML = payload;
    runInjectedScripts(box);
    if (fitOn(t)){
      watchTile(idx);
      var settle = [300, 900, 2000];
      for (var s2 = 0; s2 < settle.length; s2++){
        (function(ms){ setTimeout(function(){ autoGrowTile(idx); }, ms); })(settle[s2]);
      }
    }
    return;
  }
  if (t.kind === "custom"){
    renderCustom(box, idx, payload, (t.spec && t.spec.display) || "bar");
    if (fitOn(t)){
      watchTile(idx);
      setTimeout(function(){ autoGrowTile(idx); }, 200);
    }
    return;
  }
  // Route by SOURCE, not just kind. A cached catalog tile carries columns and
  // rows, which renderReport does not understand -- it looked for grid_data,
  // found none, and drew "Nothing to show".
  if (tileSource(t) === "catalog"){
    renderCatalogResult(t, box, idx, payload);
    return;
  }
  renderReport(t, box, idx, payload);
}

function loadWidgetTile(t, box, done, idx){
  // No preview flag: that would force NeverCache, and the cache policy the
  // widget already carries is the whole reason these tiles are cheap.
  var xhr = new XMLHttpRequest();
  xhr.open("GET", "/HomeWidgets/Embed/" + encodeURIComponent(t.widget_id), true);
  xhr.onreadystatechange = function(){
    if (xhr.readyState !== 4) return;
    var txt = xhr.responseText || "";
    if (xhr.status !== 200){
      box.innerHTML = "<span class='db-err'>Widget unavailable (HTTP "
                    + xhr.status + ").</span>";
    } else if (txt.indexOf("Error: Not authorized") >= 0){
      // Roles are enforced by TouchPoint, not by us. Say it plainly and
      // without naming what is behind it.
      box.innerHTML = "<span class='db-muted'>Not available to you.</span>";
    } else if (txt.indexOf("Error:") === 0){
      box.innerHTML = "<span class='db-err'>" + esc(txt.substring(0, 300)) + "</span>";
    } else if (!txt.replace(/\s/g, "")){
      box.innerHTML = "<span class='db-muted'>Widget returned nothing.</span>";
    } else {
      cachePut(t, txt);
      LAST_PAYLOAD[idx] = txt;
      box.innerHTML = "<div class='db-fit'></div>";
      box.firstChild.innerHTML = txt;
      runInjectedScripts(box);
      // Widget markup can settle a tick after its scripts run.
      if (fitOn(t)){
        watchTile(idx);
      }
      // The widget fetches its own data after its scripts run, and how long
      // that takes is not ours to know. Re-measure on a short ladder rather
      // than betting on a single delay.
      var settle = [300, 900, 2000, 4000];
      for (var s2 = 0; s2 < settle.length; s2++){
        (function(ms){ setTimeout(function(){ autoGrowTile(idx); }, ms); })(settle[s2]);
      }
    }
    if (done) done();
  };
  xhr.send();
}

// innerHTML never executes script elements it inserts. TouchPoint widgets ship
// their own (charts, click handlers), so each has to be recreated to run.
function runInjectedScripts(box){
  var found = box.querySelectorAll("script");
  for (var i = 0; i < found.length; i++){
    var old = found[i];
    var fresh = document.createElement("script");
    if (old.src){ fresh.src = old.src; }
    else { fresh.text = old.textContent || old.innerText || ""; }
    try { old.parentNode.replaceChild(fresh, old); }
    catch (e) { }
  }
}

// Some widgets change size after the person interacts with them: LIVESEARCH2
// is short until a search runs, then reveals results. A fixed grid cell clips
// that, so the tile grows to fit instead.
function contentHeight(fit){
  var h = Math.max(fit.offsetHeight || 0, fit.scrollHeight || 0);
  try {
    var top = fit.getBoundingClientRect().top;
    var kids = fit.querySelectorAll("*");
    for (var i = 0; i < kids.length; i++){
      var r = kids[i].getBoundingClientRect();
      // Ignore hidden nodes; a zero-height box tells us nothing.
      if (!r.height) continue;
      var bottom = r.bottom - top;
      if (bottom > h) h = bottom;
    }
  } catch (e) { }
  return h;
}

// Height mode for a tile. Charts are excluded outright: a canvas already
// fills whatever box it is given, so growing one just adds white space.
// Otherwise an explicit choice wins, and the default is to grow widgets (whose
// content is interactive and unpredictable) but not reports (whose tables are
// long by nature and would swallow the page).
// Where a tile's report should come from. An explicit source wins; otherwise
// the catalog is preferred whenever it holds that report, and Enterprise
// Reporting is the fallback. Resolved per render, so seeding the catalog fixes
// existing dashboards without reinstalling them.
var CATALOG_IDS = null;

function catalogHas(rid){
  if (CATALOG_IDS === null){
    CATALOG_IDS = {};
    var ids = (DB_BOOT && DB_BOOT.catalog_ids) || [];
    for (var i = 0; i < ids.length; i++){ CATALOG_IDS[ids[i]] = 1; }
  }
  return !!CATALOG_IDS[rid];
}

function tileSource(t){
  if (!t) return "reports";
  if (t.source === "catalog") return "catalog";
  if (t.source === "reports") return "reports";
  return catalogHas(t.report_id) ? "catalog" : "reports";
}

// A tile the viewer is not permitted to see. Deliberately quiet: no red, no
// role names, nothing that reads as a fault. It says just enough that the
// space is not mistaken for a tile that failed to load.
function restrictedHtml(){
  return "<div class='db-muted' style='display:flex;align-items:center;"
       + "justify-content:center;height:100%;min-height:60px;text-align:center;"
       + "padding:12px;font-size:12px;'>Not shown. This data is limited to "
       + "your finance team.</div>";
}

function fitOn(t){
  if (!t) return false;
  if (t.kind !== "widget" && t.display === "chart") return false;
  if (t.fit === undefined || t.fit === null) return t.kind === "widget";
  return !!t.fit;
}

function autoGrowTile(idx){
  if (!GRID) return;
  // Never while the user is dragging or resizing. This calls GRID.update on
  // the very element GridStack is mid-gesture on, which fights the mouse and
  // rewrites the node under the drag. Content taller than the tile makes it
  // fight every frame, so the tile appears glued to the resize handle. A
  // widget whose content happens to match its tile never triggers it, which
  // is why this only showed up with one church's widget.
  if (GRABBING) return;
  if (!fitOn(tiles()[idx])) return;
  var el = document.querySelector('.grid-stack-item[data-idx="' + idx + '"]');
  var body = $id("tb" + idx);
  if (!el || !body || !el.gridstackNode) return;
  if (AUTHORED_H[idx] === undefined){
    AUTHORED_H[idx] = el.gridstackNode.h || 4;
  }
  // Measure the CONTENT wrapper, which shrinks when the content does.
  var fit = body.querySelector(".db-fit");
  var content = fit ? contentHeight(fit) : 0;
  if (!content){
    content = body.scrollHeight || 0;
  }
  // Header, padding and a little slack, converted into grid rows.
  var needed = content + 52;
  var rows = Math.ceil(needed / (CELL_H + CELL_MARGIN));
  // Never below what the user chose, never past the cap: a widget that renders
  // a thousand rows should scroll, not swallow the page.
  rows = Math.max(AUTHORED_H[idx], Math.min(rows, MAX_ROWS));
  if (rows !== el.gridstackNode.h){
    try { GRID.update(el, {h: rows}); } catch (e) { }
  }
}

function watchTile(idx){
  var body = $id("tb" + idx);
  if (!body) return;
  autoGrowTile(idx);
  if (typeof MutationObserver === "undefined") return;
  var pending = null;
  var ob = new MutationObserver(function(){
    if (pending) clearTimeout(pending);
    pending = setTimeout(function(){ autoGrowTile(idx); }, 120);
  });
  ob.observe(body, {childList: true, subtree: true,
                    attributes: true, characterData: true});
  OBSERVERS.push(ob);
}

// A tile grown to fit a transient search result must not persist that height.
// Put every auto-grown tile back to its authored size before harvesting.
function restoreAuthoredHeights(){
  if (!GRID) return;
  for (var k in AUTHORED_H){
    var el = document.querySelector('.grid-stack-item[data-idx="' + k + '"]');
    if (el && el.gridstackNode && el.gridstackNode.h !== AUTHORED_H[k]){
      try { GRID.update(el, {h: AUTHORED_H[k]}); } catch (e) { }
    }
  }
}

// Enterprise Reporting speaks AG Grid (columnDefs / rowData); the catalog
// returns plain names and dicts. Normalised here so both get the same table,
// with sorting and the people actions, rather than only catalog tiles having
// them.
function normalizeGrid(gd){
  gd = gd || {};
  var defs = gd.columnDefs || gd.columns || [];
  var rows = gd.rowData || gd.rows || gd.data || [];
  var cols = [];
  for (var i = 0; i < defs.length; i++){
    cols.push(typeof defs[i] === "string"
              ? defs[i] : (defs[i].field || defs[i].headerName));
  }
  if (!cols.length && rows.length){
    for (var k in rows[0]) cols.push(k);
  }
  return {cols: cols, rows: rows};
}

function gridTable(gd){
  // Enterprise Reporting returns AG Grid's naming: columnDefs / rowData.
  // The other spellings are fallbacks in case that ever changes.
  var cols = gd.columnDefs || gd.columns || [];
  var rows = gd.rowData || gd.rows || gd.data || [];
  if (!cols.length && rows.length){
    cols = [];
    for (var k in rows[0]) cols.push({field: k, headerName: k});
  }
  if (!rows.length) return "<span class='db-muted'>No rows.</span>";
  var h = "<table class='db-t'><thead><tr>";
  for (var c = 0; c < cols.length; c++){
    h += "<th>" + esc(cols[c].headerName || cols[c].field) + "</th>";
  }
  h += "</tr></thead><tbody>";
  for (var r = 0; r < Math.min(rows.length, 200); r++){
    h += "<tr>";
    for (var c2 = 0; c2 < cols.length; c2++){
      h += "<td>" + esc(rows[r][cols[c2].field]) + "</td>";
    }
    h += "</tr>";
  }
  h += "</tbody></table>";
  if (rows.length > 200){
    h += "<div class='db-muted'>Showing 200 of " + rows.length + " rows.</div>";
  }
  return h;
}

function drawChart(box, i, cd){
  // Previewing twice reuses the same key; without this the old chart keeps
  // its canvas and its listeners.
  if (CHARTS["t" + i]){
    try { CHARTS["t" + i].destroy(); } catch (e) { }
    delete CHARTS["t" + i];
  }
  if (typeof Chart === "undefined"){
    box.innerHTML = "<span class='db-err'>Chart library did not load "
                  + "(CDN blocked?). Switch this tile to table.</span>";
    return;
  }
  box.innerHTML = "<canvas id='cv" + i + "'></canvas>";
  var ctx = $id("cv" + i).getContext("2d");
  var cfg = {
    type: cd.type || "bar",
    data: cd.data || {labels: cd.labels || [], datasets: cd.datasets || []},
    options: {responsive: true, maintainAspectRatio: false,
              plugins: {legend: {display: (cd.type === "pie" || cd.type === "doughnut")}}}
  };
  try { CHARTS["t" + i] = new Chart(ctx, cfg); }
  catch (e){ box.innerHTML = "<span class='db-err'>Chart error: " + esc(e.message) + "</span>"; }
}

// ---------- edit ----------
// "Done" used to just leave edit mode, so anything not explicitly Saved was
// lost without a word. It now offers to save what is pending.
var DIRTY = false;

function markDirty(){
  DIRTY = true;
  var b = $id("dbEditBtn");
  if (b){
    b.textContent = "Save changes";
    b.className = "btn btn-sm btn-success";
  }
}

function markClean(){
  DIRTY = false;
  var b = $id("dbEditBtn");
  if (b){
    b.textContent = EDIT ? "Done" : "Edit";
    b.className = EDIT ? "btn btn-sm btn-default" : "btn btn-sm btn-primary";
  }
}

function dbToggleEdit(){
  if (EDIT && DIRTY){ dbSave(true); return; }
  EDIT = !EDIT;
  $id("dbEditTools").style.display = EDIT ? "" : "none";
  markClean();
  renderTabs();
  renderGrid();
}

function harvestLayout(){
  if (!GRID) return;
  if (GRID_COLS < 12){
    // Coordinates here are the remapped ones. Writing them back would flatten
    // the real layout, so leave the stored positions alone.
    return;
  }
  var list = tiles();
  var nodes = null;
  try { nodes = GRID.engine ? GRID.engine.nodes : null; } catch (e) { }
  if (!nodes || !nodes.length){
    // Public API fallback. Without this, a layout change could appear to save
    // and come back in the old positions.
    try {
      var saved = GRID.save(false);
      for (var s2 = 0; s2 < saved.length; s2++){
        var idx2 = parseInt(saved[s2].id, 10);
        if (!isNaN(idx2) && list[idx2]){
          list[idx2].x = saved[s2].x; list[idx2].y = saved[s2].y;
          list[idx2].w = saved[s2].w;
          list[idx2].h = authoredH(idx2, saved[s2].h);
        }
      }
    } catch (e2) { }
    return;
  }
  for (var i = 0; i < nodes.length; i++){
    var idx = parseInt(nodes[i].el.getAttribute("data-idx"), 10);
    if (isNaN(idx) || !list[idx]) continue;
    list[idx].x = nodes[i].x;
    list[idx].y = nodes[i].y;
    list[idx].w = nodes[i].w;
    list[idx].h = authoredH(idx, nodes[i].h);
  }
}

// The height to STORE for a tile: what the user sized it to, not what
// auto-grow stretched it to on screen.
function authoredH(idx, current){
  var a = AUTHORED_H[idx];
  return (a === undefined) ? current : a;
}

function dbSave(exitEdit){
  if (GRID && GRID_COLS < 12){
    alert("This screen is too narrow to save a layout.\\n\\n"
        + "Tiles are stacked to fit, and saving now would overwrite your "
        + "arrangement with the stacked one. Open it on a wider screen.");
    return;
  }
  harvestLayout();
  var payload = JSON.stringify(DASH);
  var sp = {action: "save_dashboard", payload: encPayload(payload)};
  if (ADMIN_OWNER) sp.owner_uid = ADMIN_OWNER.uid;
  post(sp, function(r){
    if (!r.success){ alert(r.error); return; }
    DASH.id = r.id;
    DASH_LIST = null;                 // the name may have changed
    if (exitEdit){
      EDIT = false;
      $id("dbEditTools").style.display = "none";
      markClean();
      renderTabs();
      renderGrid();
      return;
    }
    markClean();
  });
}

function dbRefresh(){
  cacheClear();
  renderGrid();
}

function programKey(){
  return "tpxidb:prog:" + (DB_BOOT.user_id || 0) + ":" + ((DASH && DASH.id) || "");
}

function currentProgram(){
  try {
    var v = cacheStore().getItem(programKey());
    if (v !== null) return v;
  } catch (e) {}
  return (DASH && DASH.default_program) || "";
}

function dbSetProgram(){
  if (!DASH) return;
  var sel = $id("dbProgram");
  try { cacheStore().setItem(programKey(), sel ? sel.value : ""); } catch (e) {}
  cacheClear();
  renderGrid();
}

function fillProgramSelect(){
  var sel = $id("dbProgram");
  if (!sel) return;
  var ps = DB_BOOT.programs || [];
  var h = "<option value=''>All areas</option>";
  for (var i = 0; i < ps.length; i++){
    h += "<option value='" + esc(ps[i].id) + "'>" + esc(ps[i].name) + "</option>";
  }
  sel.innerHTML = h;
  sel.style.display = ps.length ? "" : "none";
  sel.value = currentProgram();
}

function queryKey(){
  return "tpxidb:scope:" + (DB_BOOT.user_id || 0) + ":" + ((DASH && DASH.id) || "");
}

function currentQuery(){
  try {
    var v = cacheStore().getItem(queryKey());
    if (v !== null) return v;
  } catch (e) {}
  return (DASH && DASH.default_query) || "";
}

function dbSetQuery(){
  if (!DASH) return;
  var sel = $id("dbScopeQ");
  var v = sel ? sel.value : "";
  // Kept per viewer only. Writing it onto the dashboard would turn one
  // person's choice into the default everyone else opens with.
  try { cacheStore().setItem(queryKey(), v); } catch (e) {}
  cacheClear();
  renderGrid();
}

// The dashboard's own default, as opposed to the viewer's current choice.
function fillDefQuery(){
  var sel = $id("dbDefQuery");
  if (!sel) return;
  sel.innerHTML = searchOptionsHtml("Everyone");
  sel.value = (DASH && DASH.default_query) || "";
}

function dbSetDefQuery(){
  if (!DASH) return;
  var sel = $id("dbDefQuery");
  DASH.default_query = sel ? sel.value : "";
  // Applies now for anyone who has not made their own choice on this
  // dashboard, which includes the person setting it if they never touched the
  // bar control.
  try {
    if (cacheStore().getItem(queryKey()) === null){
      var live = $id("dbScopeQ");
      if (live) live.value = DASH.default_query || "";
      cacheClear();
      renderGrid();
    }
  } catch (e) { }
  markDirty();
}

function tabCaps(){
  var list = tiles(), c = {date: 0, program: 0, people: 0, total: 0};
  for (var i = 0; i < list.length; i++){
    var cap = list[i].caps;
    if (!cap) continue;
    c.total++;
    if (cap.date) c.date++;
    if (cap.program) c.program++;
    if (cap.people) c.people++;
  }
  return c;
}

// A control is hidden when nothing on the tab can use it, and says how far it
// reaches when only some tiles can. Offering an Area filter on a giving
// dashboard was the thing that felt broken.
function applyCapsToBar(){
  var c = tabCaps();
  var specs = [["dbRange", c.date, "date range", "Each tile's own range"],
               ["dbProgram", c.program, "area", "All areas"],
               ["dbScopeQ", c.people, "search", "Everyone"]];
  for (var i = 0; i < specs.length; i++){
    var el = $id(specs[i][0]), have = specs[i][1];
    if (!el) continue;
    // Nothing on the tab can use it, or the database has nothing to offer
    // (no saved searches, no programs) -- either way it is not a choice.
    if (!c.total || !have || (el.options && el.options.length < 2)){
      el.style.display = "none";
      continue;
    }
    el.style.display = "";
    var first = el.options && el.options[0];
    if (first){
      first.text = specs[i][3]
        + (have < c.total ? (" (" + have + " of " + c.total + " tiles)") : "");
    }
    el.title = have < c.total
      ? ("Narrows " + have + " of the " + c.total + " tiles on this tab. The "
         + "rest are tagged to say why they cannot follow it.")
      : ("Narrows every tile on this tab by " + specs[i][2] + ".");
  }
}

function searchOptionsHtml(firstLabel){
  var list = DB_BOOT.searches || [];
  var mine = "", shared = "";
  for (var i = 0; i < list.length; i++){
    var o = "<option value='" + esc(list[i].name) + "'>"
          + esc(list[i].name) + "</option>";
    if (list[i].mine) mine += o; else shared += o;
  }
  return "<option value=''>" + esc(firstLabel) + "</option>"
    + (mine ? "<optgroup label='My searches'>" + mine + "</optgroup>" : "")
    + (shared ? "<optgroup label='Shared with everyone'>" + shared
                + "</optgroup>" : "");
}

function fillScopeSelect(){
  var sel = $id("dbScopeQ");
  if (!sel) return;
  var names = DB_BOOT.searches || [];
  sel.innerHTML = searchOptionsHtml("Everyone");
  // Hidden rather than shown empty: a control offering only "Everyone" reads
  // as broken.
  sel.style.display = names.length ? "" : "none";
  sel.value = currentQuery();
}

function rangeKey(){
  return "tpxidb:range:" + (DB_BOOT.user_id || 0) + ":" + ((DASH && DASH.id) || "");
}

function currentRange(){
  try {
    var v = cacheStore().getItem(rangeKey());
    if (v !== null) return v;
  } catch (e) {}
  return (DASH && DASH.date_range) || "";
}

function dbSetRange(){
  if (!DASH) return;
  var sel = $id("dbRange");
  var v = sel ? sel.value : "";
  DASH.date_range = v;
  try { cacheStore().setItem(rangeKey(), v); } catch (e) {}
  // Cached tiles were answered over the old window.
  cacheClear();
  renderGrid();
}

function dbSetCache(){
  if (!DASH) return;
  var sel = $id("dbCacheMins");
  DASH.cache_minutes = parseInt(sel.value, 10) || 0;
  var rb = $id("dbRefreshBtn");
  if (rb) rb.style.display = DASH.cache_minutes ? "" : "none";
  // Old entries were written under the previous window; clear them so the new
  // setting takes effect now rather than whenever they happened to expire.
  cacheClear();
  renderGrid();
}

function dbToggleCache(i){
  var t = tiles()[i];
  if (!t) return;
  t.nocache = !t.nocache;
  if (t.nocache){
    // Drop anything already stored for it, or turning caching off would still
    // show a stale answer until the entry expired.
    try { window.sessionStorage.removeItem(cacheKey(t)); } catch (e) { }
  }
  renderGrid();
}

function dbToggleFit(i){
  var t = tiles()[i];
  if (!t) return;
  if (t.kind !== "widget" && t.display === "chart"){
    alert("Charts already fill their tile, so height is fixed for them.");
    return;
  }
  // Store it explicitly from now on, so the default never flips underneath a
  // choice the person has made.
  t.fit = !fitOn(t);
  if (!t.fit){
    // Going fixed: drop back to the authored height rather than freezing at
    // whatever size it happened to have grown to.
    restoreAuthoredHeights();
  }
  renderGrid();
}

function dbRemoveTile(i){
  var list = tiles();
  list.splice(i, 1);
  markDirty();
  renderGrid();
}

function dbAddTab(){
  var n = prompt("Tab name", "Tab " + ((DASH.tabs || []).length + 1));
  if (!n) return;
  DASH.tabs = DASH.tabs || [];
  DASH.tabs.push({name: n, tiles: []});
  TAB = DASH.tabs.length - 1;
  markDirty();
  renderTabs();
  renderGrid();
}

function dbRenameTab(){
  var t = (DASH.tabs || [])[TAB];
  if (!t) return;
  var n = prompt("Tab name", t.name || "");
  if (!n) return;
  t.name = n;
  markDirty();
  renderTabs();
}

function dbNew(){
  var n = prompt("Dashboard name", "New Dashboard");
  if (!n) return;
  var scope = "personal";
  if (DB_BOOT.can_share && confirm("Share this dashboard with other staff?\\n\\n"
      + "OK = shared, Cancel = just for me")){
    scope = "shared";
  }
  post({action: "new_dashboard", d_name: n, d_scope: scope}, function(r){
    if (!r.success){ alert(r.error); return; }
    DASH_LIST = null;
    dbOpen(r.id);
  });
}

function dbDelete(){
  if (!DASH) return;
  if (!confirm("Delete \\"" + DASH.name + "\\"? This cannot be undone.")) return;
  var dp = {action: "delete_dashboard", d_id: DASH.id,
            d_scope: DASH.scope || "personal"};
  if (ADMIN_OWNER) dp.owner_uid = ADMIN_OWNER.uid;
  post(dp, function(r){
    if (!r.success){ alert(r.error); return; }
    DASH_LIST = null;
    dbHome();
  });
}

// ---------- add tile ----------
function dbAddTile(){
  dbModal("Add a tile",
    "<div class='db-muted' style='padding:26px 6px;text-align:center;'>"
    + "<div style='font-size:15px;margin-bottom:6px;'>Loading reports"
    + "<span id='dbLoadDots'></span></div>"
    + "<div>Fetching the report catalog from DisplayCache.</div></div>");
  loadDots(0);

  // One job per thing that actually has to be fetched, counted AFTER the list
  // is built. The previous version seeded the counter with a fixed number and
  // adjusted it as it went, so a cached source could take the count to zero
  // while another was still in flight -- and the picker rendered with, for
  // instance, "TouchPoint widgets (0)" because that fetch had not landed.
  var jobs = [];

  if (!CATALOG){
    jobs.push(function(done){
      post({action: "catalog_browse"}, function(cb){
        CATALOG = (cb && cb.success) ? cb : {reports: [], dashboards: []};
        done();
      });
    });
  }

  if (!DB_BOOT.reports_ok){
    REPORTS = [];
    // Not an error on its own: the catalog serves reports without it. Only
    // reports written inside Enterprise Reporting are actually unavailable.
    REPORTS_ERR = "Enterprise Reporting is not installed here (or is named "
                + "something this script has not been told about). Catalog "
                + "reports, widget tiles and built tiles all work without it; "
                + "only reports created inside that script are missing.";
  } else if (!REPORTS){
    jobs.push(function(done){
      postReports({action: "list_reports"}, function(r){
        REPORTS = (r && r.success) ? r.reports : [];
        if (!REPORTS.length){
          // Say WHERE the request went: "Unknown action" reads the same
          // whether the reporting script is stale or never received it.
          REPORTS_ERR = ((r && r.error) || "No response")
            + "  [posted to /PyScriptForm/" + DB_BOOT.reports_script + "]";
        }
        done();
      });
    });
  }

  if (!WIDGETS){
    jobs.push(function(done){
      post({action: "list_widgets"}, function(r){
        WIDGETS = (r && r.success) ? r.widgets : [];
        WIDGETS_SKIPPED = (r && r.skipped) || {};
        // Kept apart from "you have none": an unreachable list and an empty
        // one look identical on screen otherwise.
        WIDGETS_ERR = (r && r.success) ? "" : ((r && r.error) || "No response");
        done();
      });
    });
  }

  if (!CUSTOM_META){
    jobs.push(function(done){
      post({action: "custom_meta"}, function(r){
        CUSTOM_META = (r && r.success) ? r : {domains: [], searches: []};
        done();
      });
    });
  }

  var left = jobs.length;
  if (!left){ stopDots(); showTilePicker(); return; }
  var fire = function(){
    if (--left === 0){ stopDots(); showTilePicker(); }
  };
  for (var i = 0; i < jobs.length; i++){ jobs[i](fire); }
}

function pickSource(which){
  PICK_SRC = which;
  var ids = {reports: "dbSrcReports", widgets: "dbSrcWidgets",
             build: "dbSrcBuild", links: "dbSrcLinks", checks: "dbSrcChecks"};
  var panes = {reports: "dbPickReports", widgets: "dbPickWidgets",
               build: "dbPickBuild", links: "dbPickLinks",
               checks: "dbPickChecks"};
  for (var k in ids){
    var b = $id(ids[k]);
    if (b) b.className = "db-tab" + (which === k ? " on" : "");
    var pn = $id(panes[k]);
    if (pn) pn.style.display = which === k ? "" : "none";
  }
  if (which === "build") syncBuilder();
  if (which === "checks") loadOpsChecks();
}

var DOTS_TIMER = null;

function loadDots(i){
  var el = $id("dbLoadDots");
  if (!el) return;
  el.textContent = new Array((i % 4) + 1).join(".");
  DOTS_TIMER = setTimeout(function(){ loadDots(i + 1); }, 400);
}

function stopDots(){
  if (DOTS_TIMER){ clearTimeout(DOTS_TIMER); DOTS_TIMER = null; }
}

function showTilePicker(){
  // Counted from the SAME list the tab renders. REPORTS holds only the
  // Enterprise Reporting entries until reportPickerHtml() merges the catalog
  // in, and that runs further down this very expression -- so the label read
  // "Reports (0)" above a list of a hundred and fifty.
  var nReports = allPickerReports().length;
  var h = "<div class='db-tabs' style='margin-bottom:10px;'>"
        + "<button id='dbSrcReports' class='db-tab on' "
        + "onclick=\\"pickSource('reports')\\">Reports ("
        + nReports + ")</button>"
        + "<button id='dbSrcWidgets' class='db-tab' "
        + "onclick=\\"pickSource('widgets')\\">TouchPoint widgets ("
        + (WIDGETS ? WIDGETS.length : 0) + ")</button>"
        + "<button id='dbSrcBuild' class='db-tab' "
        + "onclick=\\"pickSource('build')\\">Build one</button>"
        + "<button id='dbSrcLinks' class='db-tab' "
        + "onclick=\\"pickSource('links')\\">Links</button>"
        + "<button id='dbSrcChecks' class='db-tab' "
        + "onclick=\\"pickSource('checks')\\">Checks</button></div>"
        + "<div id='dbPickReports'>" + reportPickerHtml() + "</div>"
        + "<div id='dbPickWidgets' style='display:none;'>" + widgetPickerHtml() + "</div>"
        + "<div id='dbPickBuild' style='display:none;'>" + builderHtml() + "</div>"
        + "<div id='dbPickLinks' style='display:none;'>" + linksEditorHtml(null) + "</div>"
        + "<div id='dbPickChecks' style='display:none;'>"
        + "<span class='db-muted'>Loading checks...</span></div>";
  dbModal("Add a tile", h);
  pickSource(PICK_SRC || "reports");
  lnkDraw();
}

// Font Awesome 4 names, which is what TouchPoint ships and what the
// QuickLinks widget uses.
// Lifted verbatim from TPxi_QuickLinksAdmin so the two offer the same icons
// under the same names. 163 across nine categories.
var FA_ICONS = {
    "General": [
        {i:"fa-home",l:"Home"},{i:"fa-church",l:"Church"},{i:"fa-globe",l:"Globe"},{i:"fa-briefcase",l:"Briefcase"},
        {i:"fa-star",l:"Star"},{i:"fa-flag",l:"Flag"},{i:"fa-heart",l:"Heart"},{i:"fa-bookmark",l:"Bookmark"},
        {i:"fa-building",l:"Building"},{i:"fa-university",l:"University"},{i:"fa-map-marker-alt",l:"Map Pin"},
        {i:"fa-calendar",l:"Calendar"},{i:"fa-calendar-alt",l:"Calendar Alt"},{i:"fa-calendar-check",l:"Calendar Check"},
        {i:"fa-clock",l:"Clock"},{i:"fa-bell",l:"Bell"},{i:"fa-bullhorn",l:"Bullhorn"},{i:"fa-gift",l:"Gift"},
        {i:"fa-trophy",l:"Trophy"},{i:"fa-child",l:"Child"},{i:"fa-baby",l:"Baby"},{i:"fa-cross",l:"Cross"},
        {i:"fa-hands-praying",l:"Praying"},{i:"fa-dove",l:"Dove"},{i:"fa-book",l:"Book"},
        {i:"fa-book-open",l:"Book Open"},{i:"fa-bible",l:"Bible"},{i:"fa-graduation-cap",l:"Grad Cap"},
        {i:"fa-bus",l:"Bus"},{i:"fa-car",l:"Car"},{i:"fa-plane",l:"Plane"},{i:"fa-anchor",l:"Anchor"}
    ],
    "People": [
        {i:"fa-user",l:"User"},{i:"fa-users",l:"Users"},{i:"fa-user-plus",l:"User Plus"},
        {i:"fa-user-minus",l:"User Minus"},{i:"fa-user-check",l:"User Check"},{i:"fa-user-lock",l:"User Lock"},
        {i:"fa-user-shield",l:"User Shield"},{i:"fa-user-cog",l:"User Cog"},{i:"fa-user-slash",l:"User Slash"},
        {i:"fa-user-tie",l:"User Tie"},{i:"fa-user-graduate",l:"Graduate"},{i:"fa-user-nurse",l:"Nurse"},
        {i:"fa-people-group",l:"Group"},{i:"fa-person",l:"Person"},{i:"fa-person-chalkboard",l:"Chalkboard"},
        {i:"fa-users-cog",l:"Users Cog"},{i:"fa-handshake",l:"Handshake"},{i:"fa-hand-holding-heart",l:"Holding Heart"},
        {i:"fa-chalkboard-teacher",l:"Teacher"}
    ],
    "Communication": [
        {i:"fa-envelope",l:"Envelope"},{i:"fa-envelope-open-text",l:"Envelope Open"},{i:"fa-paper-plane",l:"Paper Plane"},
        {i:"fa-phone",l:"Phone"},{i:"fa-comment",l:"Comment"},{i:"fa-comments",l:"Comments"},
        {i:"fa-bullseye",l:"Bullseye"},{i:"fa-at",l:"At"},{i:"fa-inbox",l:"Inbox"},
        {i:"fa-share",l:"Share"},{i:"fa-rss",l:"RSS"},{i:"fa-wifi",l:"WiFi"},
        {i:"fa-headphones",l:"Headphones"},{i:"fa-microphone",l:"Microphone"}
    ],
    "Finance": [
        {i:"fa-dollar-sign",l:"Dollar"},{i:"fa-credit-card",l:"Credit Card"},{i:"fa-wallet",l:"Wallet"},
        {i:"fa-hand-holding-usd",l:"Holding USD"},{i:"fa-coins",l:"Coins"},{i:"fa-money-bill",l:"Money Bill"},
        {i:"fa-file-invoice-dollar",l:"Invoice"},{i:"fa-cash-register",l:"Cash Register"},
        {i:"fa-building-columns",l:"Columns"},{i:"fa-receipt",l:"Receipt"},{i:"fa-piggy-bank",l:"Piggy Bank"},
        {i:"fa-chart-pie",l:"Chart Pie"},{i:"fa-calculator",l:"Calculator"}
    ],
    "Charts": [
        {i:"fa-chart-bar",l:"Chart Bar"},{i:"fa-chart-line",l:"Chart Line"},{i:"fa-chart-area",l:"Chart Area"},
        {i:"fa-table",l:"Table"},{i:"fa-database",l:"Database"},{i:"fa-gauge",l:"Gauge"},
        {i:"fa-tachometer",l:"Tachometer"},{i:"fa-signal",l:"Signal"},{i:"fa-list",l:"List"},
        {i:"fa-list-check",l:"List Check"},{i:"fa-bars",l:"Bars"},{i:"fa-project-diagram",l:"Diagram"}
    ],
    "Medical": [
        {i:"fa-hospital",l:"Hospital"},{i:"fa-first-aid",l:"First Aid"},{i:"fa-heartbeat",l:"Heartbeat"},
        {i:"fa-stethoscope",l:"Stethoscope"},{i:"fa-pills",l:"Pills"},{i:"fa-ambulance",l:"Ambulance"},
        {i:"fa-medkit",l:"Medkit"}
    ],
    "Security": [
        {i:"fa-lock",l:"Lock"},{i:"fa-unlock",l:"Unlock"},{i:"fa-shield-alt",l:"Shield"},
        {i:"fa-key",l:"Key"},{i:"fa-eye",l:"Eye"},{i:"fa-eye-slash",l:"Eye Slash"},
        {i:"fa-fire",l:"Fire"},{i:"fa-camera",l:"Camera"},{i:"fa-fingerprint",l:"Fingerprint"},
        {i:"fa-id-card",l:"ID Card"},{i:"fa-user-secret",l:"User Secret"}
    ],
    "Technology": [
        {i:"fa-cogs",l:"Cogs"},{i:"fa-cog",l:"Cog"},{i:"fa-code",l:"Code"},{i:"fa-server",l:"Server"},
        {i:"fa-tools",l:"Tools"},{i:"fa-toolbox",l:"Toolbox"},{i:"fa-plug",l:"Plug"},{i:"fa-print",l:"Print"},
        {i:"fa-tag",l:"Tag"},{i:"fa-tags",l:"Tags"},{i:"fa-desktop",l:"Desktop"},{i:"fa-laptop",l:"Laptop"},
        {i:"fa-mobile-alt",l:"Mobile"},{i:"fa-battery-full",l:"Battery"},{i:"fa-bolt",l:"Bolt"},
        {i:"fa-cloud",l:"Cloud"},{i:"fa-folder",l:"Folder"},{i:"fa-hdd",l:"HDD"},
        {i:"fa-windows",l:"Windows"},{i:"fa-network-wired",l:"Network"}
    ],
    "UI": [
        {i:"fa-check",l:"Check"},{i:"fa-check-circle",l:"Check Circle"},{i:"fa-check-square",l:"Check Square"},
        {i:"fa-plus",l:"Plus"},{i:"fa-plus-circle",l:"Plus Circle"},{i:"fa-minus",l:"Minus"},
        {i:"fa-edit",l:"Edit"},{i:"fa-pen-to-square",l:"Pen Square"},{i:"fa-trash",l:"Trash"},
        {i:"fa-download",l:"Download"},{i:"fa-upload",l:"Upload"},{i:"fa-file-alt",l:"File"},
        {i:"fa-file-export",l:"File Export"},{i:"fa-arrow-right",l:"Arrow Right"},{i:"fa-arrow-left",l:"Arrow Left"},
        {i:"fa-arrows-alt",l:"Arrows"},{i:"fa-sign-in-alt",l:"Sign In"},{i:"fa-sign-out-alt",l:"Sign Out"},
        {i:"fa-refresh",l:"Refresh"},{i:"fa-sync",l:"Sync"},{i:"fa-search",l:"Search"},
        {i:"fa-filter",l:"Filter"},{i:"fa-sort",l:"Sort"},{i:"fa-ellipsis-h",l:"Ellipsis"},
        {i:"fa-th",l:"Grid"},{i:"fa-circle-o-notch",l:"Spinner"},{i:"fa-hourglass-half",l:"Hourglass"},
        {i:"fa-question-circle",l:"Question"},{i:"fa-info-circle",l:"Info"},{i:"fa-exclamation-triangle",l:"Warning"},
        {i:"fa-layer-group",l:"Layers"},{i:"fa-cubes",l:"Cubes"},{i:"fa-tasks",l:"Tasks"},
        {i:"fa-history",l:"History"},{i:"fa-flask",l:"Flask"}
    ]
};
var LINK_ROWS = [];       // rows being edited in the links tile editor
var LINK_TILE = -1;       // index being edited, or -1 for a new tile

function linksEditorHtml(t){
  LINK_TILE = -1;
  LINK_ROWS = [];
  if (t && t.links){
    for (var i = 0; i < t.links.length; i++){
      var r = t.links[i];
      LINK_ROWS.push({label: r.label || "", url: r.url || "",
                      icon: r.icon || "fa-link", cat: r.cat || "",
                      roles: r.roles || "", newtab: !!r.newtab});
    }
  }
  if (!LINK_ROWS.length) LINK_ROWS.push({label: "", url: "", newtab: true});
  return "<p class='db-muted' style='margin-top:0;'>Shortcuts for this "
       + "dashboard: a report, an involvement, a Search Builder result, "
       + "anything with an address. Paste a full web address, or a path on "
       + "this site such as <code>/PyScript/MyReport</code>. These are saved "
       + "with the dashboard, so a copy someone else installs starts empty "
       + "rather than pointing at your pages.</p>"
       + "<div><label>Tile title</label>"
       + "<input id='lnkTitle' style='width:100%;padding:6px;' value='"
       + esc((t && t.title) || "Links") + "'></div>"
       + "<div id='lnkRows' style='margin-top:10px;'></div>"
       + "<button class='btn btn-xs btn-default' onclick='lnkAdd()'>"
       + "+ Add another</button>"
       + "<div style='margin-top:12px;'>"
       + "<button class='btn btn-sm btn-primary' onclick='lnkSave()'>"
       + "Save links tile</button></div>";
}

var ICON_OPEN = -1;      // row whose icon picker is showing
var ICON_CAT = "";       // category filter inside that picker
var ROLE_OPEN = -1;      // row whose role picker is showing

function lnkIconsFor(i){
  if (ICON_OPEN !== i) return "";
  var cats = [];
  for (var c in FA_ICONS) cats.push(c);
  var v = ((($id("icoFind") || {}).value) || "").toLowerCase();
  var h = "<div style='border:1px solid #dde3ea;border-radius:4px;padding:8px;"
        + "margin:0 0 8px;background:#fbfcfd;'>"
        + "<input id='icoFind' placeholder='Search icons...' "
        + "style='width:100%;padding:5px;margin-bottom:6px;' value='" + esc(v)
        + "' oninput='lnkDraw()'>"
        + "<div style='display:flex;flex-wrap:wrap;gap:4px;margin-bottom:6px;'>"
        + "<button class='db-chip" + (ICON_CAT ? "" : " on")
        + "' onclick='lnkIconCat(\'\')'>All</button>";
  for (var a = 0; a < cats.length; a++){
    h += "<button class='db-chip" + (ICON_CAT === cats[a] ? " on" : "")
      + "' onclick='lnkIconCat(\'" + esc(cats[a]) + "\')'>"
      + esc(cats[a]) + "</button>";
  }
  h += "</div><div style='display:grid;gap:4px;max-height:180px;overflow:auto;"
    + "grid-template-columns:repeat(auto-fill,minmax(62px,1fr));'>";
  var shown = 0;
  for (var b = 0; b < cats.length; b++){
    if (ICON_CAT && cats[b] !== ICON_CAT) continue;
    var arr = FA_ICONS[cats[b]];
    for (var k = 0; k < arr.length; k++){
      var nm = arr[k].l + " " + arr[k].i;
      if (v && nm.toLowerCase().indexOf(v) < 0) continue;
      shown++;
      h += "<span class='db-ico2" + (LINK_ROWS[i].icon === arr[k].i ? " on" : "")
        + "' data-i='" + i + "' data-ico='" + esc(arr[k].i)
        + "' onclick='lnkSetIcon(this)' title='" + esc(arr[k].i) + "'>"
        + "<i class='fa " + esc(arr[k].i) + "'></i>"
        + "<span>" + esc(arr[k].l) + "</span></span>";
    }
  }
  if (!shown) h += "<span class='db-muted'>No icon matches that.</span>";
  h += "</div><div style='margin-top:6px;'>"
    + "<input id='icoCustom' placeholder='Or type a class, e.g. fa-user-graduate' "
    + "style='width:100%;padding:5px;'>"
    + " <button class='btn btn-xs btn-default' data-i='" + i
    + "' onclick='lnkCustomIcon(this)'>Use it</button></div></div>";
  return h;
}

function lnkIconCat(c){ ICON_CAT = c; lnkDraw(); }

function lnkSetIcon(el){
  var i = parseInt(el.getAttribute("data-i"), 10);
  if (isNaN(i) || !LINK_ROWS[i]) return;
  LINK_ROWS[i].icon = el.getAttribute("data-ico");
  ICON_OPEN = -1;
  lnkDraw();
}

function lnkCustomIcon(btn){
  var i = parseInt(btn.getAttribute("data-i"), 10);
  var v = ((($id("icoCustom") || {}).value) || "").trim();
  if (!v || isNaN(i) || !LINK_ROWS[i]) return;
  LINK_ROWS[i].icon = v;
  ICON_OPEN = -1;
  lnkDraw();
}

function lnkToggleIcon(btn){
  var i = parseInt(btn.getAttribute("data-i"), 10);
  ICON_OPEN = (ICON_OPEN === i) ? -1 : i;
  ICON_CAT = "";
  lnkDraw();
}

// Roles as removable chips over a filtered list, the way the QuickLinks admin
// does it. Typing a name that is not a role would silently hide the link from
// everyone, so it is picked rather than typed.
function lnkRolesFor(i){
  var picked = String(LINK_ROWS[i].roles || "").split(",");
  var h = "<div style='display:flex;gap:4px;flex-wrap:wrap;align-items:center;"
        + "flex:1;'>";
  var any = false;
  for (var p2 = 0; p2 < picked.length; p2++){
    var rn = picked[p2].trim();
    if (!rn) continue;
    any = true;
    h += "<span class='db-roletag'>" + esc(rn)
      + "<span data-i='" + i + "' data-r='" + esc(rn)
      + "' onclick='lnkDropRole(this)' title='Remove'>&times;</span></span>";
  }
  if (!any){
    h += "<span class='db-muted' style='font-size:11px;'>Everyone</span>";
  }
  h += "<button class='btn btn-xs btn-default' data-i='" + i
    + "' onclick='lnkToggleRoles(this)'>Roles</button></div>";
  if (ROLE_OPEN !== i) return h;
  var v = ((($id("rolFind") || {}).value) || "").toLowerCase();
  h += "<div style='border:1px solid #dde3ea;border-radius:4px;padding:8px;"
    + "margin:4px 0 8px;background:#fbfcfd;width:100%;'>"
    + "<input id='rolFind' placeholder='Filter roles...' "
    + "style='width:100%;padding:5px;margin-bottom:6px;' value='" + esc(v)
    + "' oninput='lnkDraw()'>"
    + "<div style='max-height:150px;overflow:auto;display:grid;gap:2px;"
    + "grid-template-columns:repeat(auto-fill,minmax(150px,1fr));'>";
  var all = ROLES || [];
  for (var r2 = 0; r2 < all.length; r2++){
    if (v && all[r2].toLowerCase().indexOf(v) < 0) continue;
    var on = false;
    for (var q2 = 0; q2 < picked.length; q2++){
      if (picked[q2].trim() === all[r2]) on = true;
    }
    h += "<label style='font-weight:normal;font-size:12px;'>"
      + "<input type='checkbox' data-i='" + i + "' data-r='" + esc(all[r2]) + "'"
      + (on ? " checked" : "") + " onclick='lnkPickRole(this)'> "
      + esc(all[r2]) + "</label>";
  }
  h += "</div></div>";
  return h;
}

function lnkToggleRoles(btn){
  var i = parseInt(btn.getAttribute("data-i"), 10);
  ROLE_OPEN = (ROLE_OPEN === i) ? -1 : i;
  if (ROLE_OPEN >= 0 && !ROLES){
    post({action: "list_roles"}, function(r){
      ROLES = (r && r.success) ? r.roles : [];
      lnkDraw();
    });
    return;
  }
  lnkDraw();
}

function lnkPickRole(cb){
  var i = parseInt(cb.getAttribute("data-i"), 10);
  var rn = cb.getAttribute("data-r");
  if (isNaN(i) || !LINK_ROWS[i]) return;
  var cur = [];
  var parts = String(LINK_ROWS[i].roles || "").split(",");
  for (var a = 0; a < parts.length; a++){
    var t = parts[a].trim();
    if (t && t !== rn) cur.push(t);
  }
  if (cb.checked) cur.push(rn);
  LINK_ROWS[i].roles = cur.join(",");
  lnkDraw();
}

function lnkDropRole(el){
  var i = parseInt(el.getAttribute("data-i"), 10);
  var rn = el.getAttribute("data-r");
  if (isNaN(i) || !LINK_ROWS[i]) return;
  var cur = [], parts = String(LINK_ROWS[i].roles || "").split(",");
  for (var a = 0; a < parts.length; a++){
    var t = parts[a].trim();
    if (t && t !== rn) cur.push(t);
  }
  LINK_ROWS[i].roles = cur.join(",");
  lnkDraw();
}

function lnkDraw(){
  var box = $id("lnkRows");
  if (!box) return;
  var h = "";
  for (var i = 0; i < LINK_ROWS.length; i++){
    var r = LINK_ROWS[i];
    h += "<div style='display:flex;gap:6px;margin-bottom:5px;"
      + "align-items:center;'>"
      + "<input data-i='" + i + "' data-fld='label' oninput='lnkEdit(this)' "
      + "placeholder='What to call it' style='flex:1;padding:5px;' value='"
      + esc(r.label) + "'>"
      + "<input data-i='" + i + "' data-fld='url' oninput='lnkEdit(this)' "
      + "placeholder='/PyScript/Name or https://...' "
      + "style='flex:2;padding:5px;' value='" + esc(r.url) + "'>"
      + "<label class='db-muted' style='white-space:nowrap;font-weight:normal;'>"
      + "<input type='checkbox' data-i='" + i + "' data-fld='newtab' "
      + "onchange='lnkEdit(this)'" + (r.newtab ? " checked" : "")
      + "> new tab</label>"
      + "<button class='btn btn-xs btn-default' data-i='" + i + "' "
      + "onclick='lnkMove(this,-1)'" + (i === 0 ? " disabled" : "")
      + " title='Move up'>&uarr;</button>"
      + "<button class='btn btn-xs btn-default' data-i='" + i + "' "
      + "onclick='lnkMove(this,1)'"
      + (i === LINK_ROWS.length - 1 ? " disabled" : "")
      + " title='Move down'>&darr;</button>"
      + "<button class='btn btn-xs btn-default' data-i='" + i + "' "
      + "onclick='lnkDrop(this)'>&times;</button></div>"
      // The icon opens a picker and roles come from the real list, the way
      // the QuickLinks admin does both.
      + "<div style='display:flex;gap:6px;margin-bottom:3px;flex-wrap:wrap;"
      + "align-items:center;'>"
      + "<input data-i='" + i + "' data-fld='cat' oninput='lnkEdit(this)' "
      + "placeholder='Group (optional)' "
      + "style='flex:1;min-width:130px;padding:4px;font-size:12px;' value='"
      + esc(r.cat || "") + "'>"
      + "<span class='db-ico' data-i='" + i + "' onclick='lnkToggleIcon(this)'>"
      + "<i class='fa " + esc(r.icon || "fa-link") + "'></i>"
      + esc(r.icon || "fa-link") + "</span>"
      + lnkRolesFor(i)
      + "</div>"
      + lnkIconsFor(i);
  }
  box.innerHTML = h;
}

function lnkEdit(el){
  var i = parseInt(el.getAttribute("data-i"), 10);
  var f = el.getAttribute("data-fld");
  if (isNaN(i) || !LINK_ROWS[i]) return;
  LINK_ROWS[i][f] = (f === "newtab") ? el.checked : el.value;
}

function lnkAdd(){
  LINK_ROWS.push({label: "", url: "", newtab: true});
  lnkDraw();
}

function lnkMove(btn, by){
  var i = parseInt(btn.getAttribute("data-i"), 10);
  var j = i + by;
  if (isNaN(i) || j < 0 || j >= LINK_ROWS.length) return;
  var tmp = LINK_ROWS[i];
  LINK_ROWS[i] = LINK_ROWS[j];
  LINK_ROWS[j] = tmp;
  lnkDraw();
}

function lnkDrop(btn){
  var i = parseInt(btn.getAttribute("data-i"), 10);
  if (!isNaN(i)) LINK_ROWS.splice(i, 1);
  if (!LINK_ROWS.length) LINK_ROWS.push({label: "", url: "", newtab: true});
  lnkDraw();
}

function lnkSave(){
  var keep = [], bad = [];
  for (var i = 0; i < LINK_ROWS.length; i++){
    var r = LINK_ROWS[i];
    if (!String(r.url || "").trim() && !String(r.label || "").trim()) continue;
    var u = safeUrl(r.url);
    if (!u){ bad.push(r.label || r.url); continue; }
    keep.push({label: String(r.label || "").trim() || u, url: u,
               icon: r.icon || "fa-link",
               cat: String(r.cat || "").trim(),
               roles: String(r.roles || "").trim(),
               newtab: !!r.newtab});
  }
  if (bad.length){
    alert("Left out, because the address was not understood: "
          + bad.join(", ")
          + ". Use a full web address, or a path starting with a slash.");
  }
  if (!keep.length){ alert("Add at least one link."); return; }
  var title = (($id("lnkTitle") || {}).value || "Links").trim() || "Links";
  var list = tiles();
  if (LINK_TILE >= 0 && list[LINK_TILE]){
    list[LINK_TILE].title = title;
    list[LINK_TILE].links = keep;
  } else {
    var maxY = 0;
    for (var j = 0; j < list.length; j++){
      maxY = Math.max(maxY, (list[j].y || 0) + (list[j].h || 4));
    }
    list.push({kind: "links", title: title, links: keep,
               x: 0, y: maxY, w: 4, h: 3, fit: true});
  }
  LINK_TILE = -1;
  markDirty();
  dbCloseModal();
  renderGrid();
}

// The operational checks published on DisplayCache. Only the automatic ones
// arrive here: a manual check has no SQL and belongs on a checklist, not a
// dashboard tile.
var OPS_CHECKS = null;

function loadOpsChecks(){
  var box = $id("dbPickChecks");
  if (!box || OPS_CHECKS){ if (OPS_CHECKS) drawOpsChecks(); return; }
  post({action: "ops_list"}, function(r){
    if (!r || !r.success){
      box.innerHTML = "<p class='db-err'>" + esc((r && r.error) || "No response")
        + "</p><p class='db-muted'>These come from the marketplace on "
        + "scripts.displaycache.com. Everything else keeps working without "
        + "them.</p>";
      return;
    }
    OPS_CHECKS = r.checks || [];
    drawOpsChecks();
  });
}

function drawOpsChecks(){
  var box = $id("dbPickChecks");
  if (!box) return;
  if (!OPS_CHECKS.length){
    box.innerHTML = "<span class='db-muted'>No runnable checks published.</span>";
    return;
  }
  var groups = {}, order = [];
  for (var i = 0; i < OPS_CHECKS.length; i++){
    var c = OPS_CHECKS[i].cat || "General";
    if (!groups[c]){ groups[c] = []; order.push(c); }
    groups[c].push(OPS_CHECKS[i]);
  }
  order.sort();
  var h = "<p class='db-muted'>Best-practice checks from the DisplayCache "
        + "marketplace. Each one is a list of what needs fixing, with the "
        + "steps to fix it. They run as they are published, so they take no "
        + "date range, area or search.</p>"
        + "<input id='opsFind' placeholder='Filter checks...' "
        + "style='width:100%;padding:6px 10px;margin-bottom:8px;' "
        + "oninput='opsFilter()'>"
        + "<div id='opsList' style='max-height:420px;overflow:auto;'>";
  for (var g = 0; g < order.length; g++){
    h += "<div class='opsCat'><div style='font-weight:600;margin:10px 0 4px;'>"
      + esc(order[g]) + "</div>";
    var arr = groups[order[g]];
    for (var k = 0; k < arr.length; k++){
      var c2 = arr[k];
      h += "<div class='opsRow' data-name='"
        + esc((c2.name + " " + c2.note + " " + c2.cat).toLowerCase())
        + "' style='padding:5px 8px;border-bottom:1px solid #f2f5f8;"
        + "display:flex;gap:8px;align-items:center;'>"
        + "<div style='flex:1;'><b>" + esc(c2.name) + "</b>"
        + (c2.freq ? (" <span class='db-muted' style='font-size:11px;'>"
                      + esc(c2.freq) + "</span>") : "")
        + "<div class='db-muted'>" + esc(c2.note || "") + "</div></div>"
        + "<select id='opsd_" + esc(c2.id) + "' style='padding:3px;'>"
        + "<option value='table'>table</option>"
        + "<option value='kpi'>kpi</option></select>"
        + "<button class='btn btn-xs btn-primary' data-c='" + esc(c2.id)
        + "' onclick='opsAdd(this)'>Add</button></div>";
    }
    h += "</div>";
  }
  h += "</div>";
  box.innerHTML = h;
}

function opsFilter(){
  var v = (($id("opsFind") || {}).value || "").toLowerCase();
  var rows = document.querySelectorAll("#opsList .opsRow");
  for (var i = 0; i < rows.length; i++){
    var nm = rows[i].getAttribute("data-name") || "";
    rows[i].style.display = (!v || nm.indexOf(v) >= 0) ? "" : "none";
  }
  var cats = document.querySelectorAll("#opsList .opsCat");
  for (var g = 0; g < cats.length; g++){
    var rr = cats[g].querySelectorAll(".opsRow"), any = false;
    for (var r = 0; r < rr.length; r++){
      if (rr[r].style.display !== "none") any = true;
    }
    cats[g].style.display = any ? "" : "none";
  }
}

function opsAdd(btn){
  var cid = btn.getAttribute("data-c");
  var disp = (($id("opsd_" + cid) || {}).value) || "table";
  var name = cid;
  for (var i = 0; i < (OPS_CHECKS || []).length; i++){
    if (OPS_CHECKS[i].id === cid) name = OPS_CHECKS[i].name;
  }
  var list = tiles(), maxY = 0;
  for (var j = 0; j < list.length; j++){
    maxY = Math.max(maxY, (list[j].y || 0) + (list[j].h || 4));
  }
  list.push({kind: "opscheck", check_id: cid, title: name, display: disp,
             x: 0, y: maxY, w: 4, h: (disp === "kpi" ? 3 : 6),
             fit: (disp === "kpi")});
  markDirty();
  dbCloseModal();
  renderGrid();
}

function loadOpsTile(t, box, done, idx){
  post({action: "ops_run", check_id: t.check_id}, function(r){
    if (!r || !r.success){
      box.innerHTML = "<span class='db-err'>"
        + esc((r && r.error) || "No response") + "</span>";
      if (done) done();
      return;
    }
    cachePut(t, r);
    LAST_PAYLOAD[idx] = r;
    scopeLabel(idx, r);
    renderOpsResult(t, box, idx, r);
    if (done) done();
  });
}

function renderOpsResult(t, box, i, r){
  var rows = r.rows || [];
  // A check finding nothing is the good outcome, so it says so rather than
  // reading like a tile that failed to load.
  if (!rows.length){
    box.innerHTML = "<div class='db-fit'><div class='db-kpi' "
      + "style='color:#1b6b33;'>0</div><div class='db-muted'>"
      + "Nothing to fix</div></div>";
    return;
  }
  if ((t.display || "table") === "kpi"){
    var th = parseInt(r.threshold, 10);
    var over = !isNaN(th) && th > 0 && rows.length > th;
    box.innerHTML = "<div class='db-fit'><div class='db-kpi'"
      + (over ? " style='color:#c0392b;'" : "") + ">"
      + esc(fmtNum(rows.length)) + esc(r.truncated ? "+" : "") + "</div>"
      + "<div class='db-muted'>to review</div></div>";
    return;
  }
  box.innerHTML = "<div class='db-fit'></div>";
  TABLE_STATE[i] = {cols: r.columns || [], rows: rows,
                    sort: (TABLE_STATE[i] || {}).sort, box: box, tile: t};
  drawCatalogTable(i);
  if (fitOn(t)){
    watchTile(i);
    setTimeout(function(){ autoGrowTile(i); }, 200);
  }
}

function widgetPickerHtml(){
  if (!WIDGETS || !WIDGETS.length){
    if (WIDGETS_ERR){
      return "<p class='db-err'>Could not read the widget list: "
           + esc(WIDGETS_ERR) + "</p><p class='db-muted'>This is not the same "
           + "as having none. Close this and try again.</p>";
    }
    return "<p class='db-muted'>No TouchPoint widgets are available to you. "
         + "They must be enabled under Admin, and carry at least one role you "
         + "hold. A widget with no roles cannot be embedded at all.</p>";
  }
  var h = "<p class='db-muted'>These render through TouchPoint and honor the "
        + "cache and roles already set on them. Ones marked "
        + "<i>(off home page)</i> are switched off on TouchPoint&#39;s own home "
        + "page but work perfectly well as tiles.</p>";
  var sk = WIDGETS_SKIPPED || {};
  if (sk.no_roles || sk.not_mine){
    h += "<p class='db-muted'>";
    if (sk.no_roles){
      h += "<b>" + sk.no_roles + "</b> widget" + (sk.no_roles === 1 ? " is" : "s are")
        + " not listed because no role is assigned to "
        + (sk.no_roles === 1 ? "it" : "them")
        + ". TouchPoint will not embed a widget with no roles, even though its "
        + "home page shows it. Give it a role under Admin to use it here. ";
    }
    if (sk.not_mine){
      h += "<b>" + sk.not_mine + "</b> more need a role you do not hold.";
    }
    h += "</p>";
  }
  for (var i = 0; i < WIDGETS.length; i++){
    var w = WIDGETS[i];
    h += "<div style='padding:6px 8px;border-bottom:1px solid #f2f5f8;"
      + "display:flex;gap:8px;align-items:center;'>"
      + "<div style='flex:1;'><b>" + esc(w.name) + "</b>"
      + (w.enabled ? "" : " <span class='db-muted' title='Switched off on "
                        + "TouchPoint&#39;s home page. It still works here.'>"
                        + "(off home page)</span>")
      + "<div class='db-muted'>" + esc(w.description || "")
      + (w.cache_hours ? (" &middot; cached " + w.cache_hours + "h") : "")
      + " &middot; " + esc(w.roles) + "</div></div>"
      + "<button class='btn btn-xs btn-primary' onclick='dbPickWidget("
      + w.id + ")'>Add</button></div>";
  }
  return h;
}

function builderHtml(){
  var m = CUSTOM_META || {domains: [], searches: []};
  if (!m.domains.length){
    return "<p class='db-err'>Could not load the builder catalog.</p>";
  }
  var h = "<p class='db-muted'>Start from a Search Builder search to choose WHO, "
        + "then pick what to measure about them. Leave the search blank for "
        + "everyone.</p><div style='display:grid;"
        + "grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;'>";

  h += "<div><label>Search Builder search</label>"
     + "<select id='bldQuery' style='width:100%;padding:5px;'>"
     + "<option value=''>-- everyone --</option>";
  for (var i = 0; i < m.searches.length; i++){
    h += "<option value='" + esc(m.searches[i].name) + "'>"
      + esc(m.searches[i].name) + "</option>";
  }
  h += "</select></div>";

  h += "<div><label>Data</label><select id='bldDomain' style='width:100%;padding:5px;' "
     + "onchange='syncBuilder()'>";
  for (var d = 0; d < m.domains.length; d++){
    h += "<option value='" + esc(m.domains[d].key) + "'>"
      + esc(m.domains[d].label) + "</option>";
  }
  h += "</select></div>";

  h += "<div><label>Measure</label>"
     + "<select id='bldMeasure' style='width:100%;padding:5px;'></select></div>";
  h += "<div><label>Group by</label>"
     + "<select id='bldDim' style='width:100%;padding:5px;' onchange='syncBuilder()'>"
     + "</select></div>";
  h += "<div id='bldMonthsWrap'><label>Time span</label>"
     + "<select id='bldMonths' style='width:100%;padding:5px;'>"
     + "<option value='3'>Last 3 months</option>"
     + "<option value='6'>Last 6 months</option>"
     + "<option value='12' selected>Last 12 months</option>"
     + "<option value='24'>Last 2 years</option>"
     + "<option value='60'>Last 5 years</option></select></div>";
  h += "<div><label>Show as</label>"
     + "<select id='bldDisplay' style='width:100%;padding:5px;'>"
     + "<option value='bar'>Bar chart</option>"
     + "<option value='line'>Line chart</option>"
     + "<option value='pie'>Pie chart</option>"
     + "<option value='doughnut'>Donut</option>"
     + "<option value='table'>Table</option>"
     + "<option value='number'>Single number</option></select></div>";
  h += "<div><label>Title</label>"
     + "<input id='bldTitle' style='width:100%;padding:5px;' placeholder='optional'>"
     + "</div>";
  var fl = m.filters || [];
  for (var f2 = 0; f2 < fl.length; f2++){
    h += "<div><label>" + esc(fl[f2].label) + "</label>"
      + "<select id='bldf_" + esc(fl[f2].key) + "' multiple size='4' "
      + "style='width:100%;padding:5px;'>";
    for (var o2 = 0; o2 < fl[f2].options.length; o2++){
      h += "<option value='" + esc(fl[f2].options[o2].v) + "'>"
        + esc(fl[f2].options[o2].n) + "</option>";
    }
    h += "</select><div class='db-muted' style='font-size:11px;'>"
      + "Nothing selected means all.</div></div>";
  }
  h += "</div>";
  h += "<div style='margin-top:12px;'>"
     + "<button class='btn btn-sm btn-default' onclick='builderPreview()'>Preview</button>"
     + " <button class='btn btn-sm btn-primary' onclick='builderAdd()'>Add tile</button>"
     + "</div><div id='bldPreview' style='margin-top:12px;'></div>";
  return h;
}

function builderDomain(){
  var m = CUSTOM_META || {domains: []};
  var key = ($id("bldDomain") || {}).value;
  for (var i = 0; i < m.domains.length; i++){
    if (m.domains[i].key === key) return m.domains[i];
  }
  return m.domains[0];
}

// Measures and groupings belong to a data type, so they are rebuilt whenever it
// changes. Without this you can ask for "total given" grouped by "gender" on
// attendance and get a server-side rejection instead of a chart.
function syncBuilder(){
  var d = builderDomain();
  if (!d) return;
  var ms = $id("bldMeasure"), dm = $id("bldDim"), mw = $id("bldMonthsWrap");
  if (ms){
    var keep = ms.value, h = "";
    for (var i = 0; i < d.measures.length; i++){
      h += "<option value='" + esc(d.measures[i].key) + "'>"
        + esc(d.measures[i].label) + "</option>";
    }
    ms.innerHTML = h;
    ms.value = keep;
    if (!ms.value && d.measures.length) ms.value = d.measures[0].key;
  }
  if (dm){
    var keepd = dm.value, h2 = "";
    for (var j = 0; j < d.dimensions.length; j++){
      h2 += "<option value='" + esc(d.dimensions[j].key) + "'>"
         + esc(d.dimensions[j].label) + "</option>";
    }
    dm.innerHTML = h2;
    dm.value = keepd;
    if (!dm.value) dm.value = "none";
  }
  if (mw) mw.style.display = d.dates ? "" : "none";
  // "One total" is a single number; anything else needs a shape.
  var disp = $id("bldDisplay");
  if (disp && dm){
    if (dm.value === "none"){ disp.value = "number"; }
    else if (disp.value === "number"){ disp.value = "bar"; }
  }
}

function builderTitle(){
  var typed = (($id("bldTitle") || {}).value || "").trim();
  if (typed) return typed;
  // Fall back to something meaningful rather than an untitled tile.
  var ms = $id("bldMeasure");
  var mlabel = (ms && ms.options && ms.options.length)
             ? ms.options[ms.selectedIndex].text : "";
  var qv = ($id("bldQuery") || {}).value || "";
  return (mlabel + (qv ? (" - " + qv) : "")) || "Untitled tile";
}

function builderSpec(){
  var d = builderDomain();
  var dm = ($id("bldDim") || {}).value || "none";
  var out = {
    query: ($id("bldQuery") || {}).value || "",
    domain: d ? d.key : "",
    measure: ($id("bldMeasure") || {}).value || "",
    dimension: dm,
    months: parseInt(($id("bldMonths") || {}).value || "12", 10),
    display: ($id("bldDisplay") || {}).value || "bar"
  };
  // Multi-selects, so several campuses or statuses can be combined.
  var fl = (CUSTOM_META && CUSTOM_META.filters) || [];
  for (var i = 0; i < fl.length; i++){
    var el = $id("bldf_" + fl[i].key);
    if (!el || !el.options) continue;
    var picked = [];
    for (var o = 0; o < el.options.length; o++){
      if (el.options[o].selected && el.options[o].value){
        picked.push(el.options[o].value);
      }
    }
    if (picked.length) out["f_" + fl[i].key] = picked.join(",");
  }
  return out;
}

function builderPreview(){
  var spec = builderSpec();
  var out = $id("bldPreview");
  out.style.height = "";
  out.innerHTML = "<span class='db-muted'>Running...</span>";
  post({action: "run_custom", spec: encPayload(JSON.stringify(spec))}, function(r){
    if (!r || !r.success){
      out.style.height = "";
      out.innerHTML = "<span class='db-err'>" + esc((r && r.error) || "No response")
                    + "</span>";
      return;
    }
    // Draw the tile as it will actually appear, chrome and all, so the title
    // and proportions are visible before committing to it.
    var isChart = !r.single && spec.display !== "number" && spec.display !== "table";
    out.style.height = "";
    out.innerHTML = "<div class='db-muted' style='margin-bottom:6px;'>Preview</div>"
      + "<div class='db-tile' style='height:" + (isChart ? "320px" : "auto")
      + ";max-width:520px;'>"
      + "<div class='db-tile-hd' style='cursor:default;'><span>"
      + esc(builderTitle()) + "</span></div>"
      + "<div class='db-tile-bd' id='bldPreviewBody'></div></div>";
    renderCustom($id("bldPreviewBody"), "preview", r, spec.display);
  });
}

function builderAdd(){
  var spec = builderSpec();
  if (!spec.measure){ alert("Pick a measure."); return; }
  var title = builderTitle();
  var list = tiles();
  var maxY = 0;
  for (var i = 0; i < list.length; i++){
    maxY = Math.max(maxY, (list[i].y || 0) + (list[i].h || 4));
  }
  list.push({kind: "custom", title: title, spec: spec,
             x: 0, y: maxY, w: 4, h: 4,
             fit: (spec.display === "table" || spec.display === "number")});
  markDirty();
  dbCloseModal();
  renderGrid();
}

function dbPickWidget(id){
  var w = null;
  for (var i = 0; i < WIDGETS.length; i++){
    if (WIDGETS[i].id === id){ w = WIDGETS[i]; break; }
  }
  if (!w) return;
  var list = tiles();
  var maxY = 0;
  for (var j = 0; j < list.length; j++){
    maxY = Math.max(maxY, (list[j].y || 0) + (list[j].h || 4));
  }
  list.push({kind: "widget", widget_id: w.id, title: w.name,
             x: 0, y: maxY, w: 4, h: 4, fit: true});
  markDirty();
  dbCloseModal();
  renderGrid();
}

function allPickerReports(){
  // Catalog first: it is the source that works on its own. Enterprise
  // Reporting entries are appended and marked, and an id present in both
  // shows once, from the catalog.
  var out = [], seen = {};
  var cat = (CATALOG && CATALOG.reports) || [];
  for (var i = 0; i < cat.length; i++){
    var c = cat[i];
    out.push({id: c.id, name: c.name, description: c.description,
              category: c.category, scopeable: c.scopeable,
              display_types: (c.display && c.display.types
                              && c.display.types.length)
                             ? c.display.types : ["table"],
              default_display: (c.display && c.display.default) || "table",
              columns: c.columns || [],
              chart_cfg: c.display || {},
              settings: c.settings || [],
              uses_serving: (c.tokens || []).join(",").indexOf("_types") >= 0,
              installed: !!c.installed,
              source: "catalog"});
    seen[c.id] = 1;
  }
  for (var j = 0; j < (REPORTS || []).length; j++){
    if (seen[REPORTS[j].id]) continue;
    var e = {}; for (var k in REPORTS[j]){ e[k] = REPORTS[j][k]; }
    e.source = "reports";
    out.push(e);
  }
  return out;
}

function reportPickerHtml(){
  var all = allPickerReports();
  if (!all.length){
    var avail = (CATALOG && CATALOG.reports || []).length;
    return "<p class='db-err'>" + esc(REPORTS_ERR || "No reports available.")
      + "</p>" + (avail ? ("<p class='db-muted'>" + avail + " reports are "
      + "available from the catalog but none are installed. Use "
      + "<b>Library &rarr; Catalog</b> to install some.</p>") : "");
  }
  REPORTS = all;
  // The markup below renders "All" as the selected chip, so the state has to
  // start there too or a reopened picker filters by last time's category with
  // nothing highlighted.
  PICK_CAT = "";
  var cats = {};
  for (var i = 0; i < REPORTS.length; i++){
    var c = REPORTS[i].category || "other";
    (cats[c] = cats[c] || []).push(REPORTS[i]);
  }
  // Category list for the chips, taken from the grouped map above. It was a
  // second "var cats" before, which overwrote the grouping and left every
  // category heading empty.
  var catKeys = [];
  for (var ck in cats) catKeys.push(ck);
  catKeys.sort();
  var h = "<input id='dbFind' placeholder='Search reports by name...' "
    + "style='width:100%;padding:6px 10px;margin-bottom:8px;' "
    + "oninput='filterPicker()'>"
    + "<div id='dbCatChips' style='display:flex;flex-wrap:wrap;gap:5px;"
    + "margin-bottom:10px;'>"
    + "<button class='db-chip on' data-cat='' onclick='pickCat(this)'>All ("
    + all.length + ")</button>";
  for (var c1 = 0; c1 < catKeys.length; c1++){
    h += "<button class='db-chip' data-cat='" + esc(catKeys[c1])
      + "' onclick='pickCat(this)'>"
      + esc(catKeys[c1].replace(/_/g, " ")) + " ("
      + cats[catKeys[c1]].length + ")</button>";
  }
  h += "</div>";
  h += "<div id='dbPickList' style='max-height:460px;overflow:auto;'>";
  var keys = [];
  for (var k in cats) keys.push(k);
  keys.sort();
  for (var j = 0; j < keys.length; j++){
    h += "<div class='db-pick-cat' data-cat='" + esc(keys[j]) + "'>"
      + "<div style='font-weight:600;margin:10px 0 4px;text-transform:capitalize;'>"
      + esc(keys[j].replace(/_/g, " ")) + "</div>";
    var arr = cats[keys[j]];
    for (var m = 0; m < arr.length; m++){
      var rp = arr[m];
      h += "<div class='db-pick' data-cat='" + esc(rp.category || "other")
        + "' data-name='" + esc((rp.name + " " + rp.description).toLowerCase())
        + "' style='padding:5px 8px;border-bottom:1px solid #f2f5f8;display:flex;gap:8px;align-items:center;'>"
        + "<div style='flex:1;'><b>" + esc(rp.name) + "</b>"
        + (rp.scopeable ? " <span class='db-muted' style='font-size:11px;'>"
                        + "scopeable</span>" : "")
        + (rp.source === "catalog" && !rp.installed
           ? " <span style='font-size:11px;color:#8a5a00;'>from catalog</span>" : "")
        + "<div class='db-muted'>"
        + esc(rp.description || "") + "</div></div>"
        + "<select id='disp_" + esc(rp.id) + "' style='padding:3px;'>";
      var types = rp.display_types || ["table"];
      for (var d = 0; d < types.length; d++){
        h += "<option value='" + esc(types[d]) + "'"
          + (types[d] === rp.default_display ? " selected" : "") + ">"
          + esc(types[d]) + "</option>";
      }
      h += "</select><button class='btn btn-xs btn-primary' onclick=\\"dbPickOpen('"
        + esc(rp.id) + "','" + esc(rp.name).replace(/'/g, "")
        + "')\\">Add</button></div>";
    }
    h += "</div>";
  }
  h += "</div>";
  return h;
}

function pickCat(btn){
  var wrap = $id("dbCatChips");
  var all = wrap ? wrap.querySelectorAll(".db-chip") : [];
  for (var i = 0; i < all.length; i++){ all[i].className = "db-chip"; }
  btn.className = "db-chip on";
  PICK_CAT = btn.getAttribute("data-cat") || "";
  filterPicker();
  // Back to the top: after narrowing to one category, staying scrolled into
  // the middle of the old list looks like nothing happened.
  var list = $id("dbPickList");
  if (list) list.scrollTop = 0;
}

function filterPicker(){
  var v = (($id("dbFind") || {}).value || "").toLowerCase();
  var c = PICK_CAT || "";
  var items = document.querySelectorAll("#dbPickList .db-pick");
  for (var i = 0; i < items.length; i++){
    var nm = items[i].getAttribute("data-name") || "";
    var ct = items[i].getAttribute("data-cat") || "";
    var show = (!v || nm.indexOf(v) >= 0) && (!c || ct === c);
    items[i].style.display = show ? "" : "none";
  }
  // Hide a category heading whose rows are all filtered away.
  var groups = document.querySelectorAll("#dbPickList .db-pick-cat");
  for (var g = 0; g < groups.length; g++){
    var rows = groups[g].querySelectorAll(".db-pick");
    var any = false;
    for (var r = 0; r < rows.length; r++){
      if (rows[r].style.display !== "none") any = true;
    }
    groups[g].style.display = any ? "" : "none";
  }
}

// Charts and KPIs need to know what to group by and what to measure. A
// catalog report is a plain result set, so unless its definition already says
// (chart_label_col / chart_data_cols) the tile has to be told here.
function cfgDispChanged(){
  var box = $id("cfgAgg");
  if (!box) return;
  var disp = (($id("cfgDisp") || {}).value) || "table";
  if (disp === "table"){
    box.innerHTML = "";
    return;
  }
  var cols = CFG.columns || [];
  var cc = CFG.chartCfg || {};
  var defLabel = cc.chart_label_col || cols[0] || "";
  var defVal = (cc.chart_data_cols && cc.chart_data_cols[0]) || "";
  // A definition that names its own value column is already aggregated by the
  // query, so measuring it again would be wrong: show it as-is.
  var defFn = defVal ? "none" : "count";

  var h = "<div style='border-top:1px solid #e5e5e5;padding-top:10px;'>";
  if (disp === "chart"){
    h += "<div style='display:grid;grid-template-columns:1fr 1fr;gap:10px;'>"
      + "<div><label>Group by</label><select id='aggLabel' "
      + "style='width:100%;padding:5px;'>" + colOptions(cols, defLabel)
      + "</select></div>"
      + "<div><label>Chart type</label><select id='aggChart' "
      + "style='width:100%;padding:5px;'>"
      + typeOptions(["bar", "line", "pie", "doughnut"],
                    cc.chart_type || "bar")
      + "</select></div></div>";
  }
  h += "<div style='display:grid;grid-template-columns:1fr 1fr;gap:10px;"
    + "margin-top:10px;'>"
    + "<div><label>Measure</label><select id='aggFn' style='width:100%;"
    + "padding:5px;' onchange='cfgFnChanged()'>"
    + typeOptions(["count", "sum", "avg", "min", "max", "none"], defFn)
    + "</select></div>"
    + "<div id='aggValWrap'><label>Of column</label><select id='aggVal' "
    + "style='width:100%;padding:5px;'>" + colOptions(cols, defVal)
    + "</select></div></div>";
  if (disp === "chart"){
    h += "<div style='margin-top:10px;'><label>Show at most</label>"
      + "<select id='aggTop' style='width:100%;padding:5px;'>"
      + typeOptions(["10", "15", "25", "50", "all"], "15")
      + "</select><div class='db-muted'>Largest first. A chart of hundreds of "
      + "bars is unreadable, so the rest are grouped as Other.</div></div>";
  }
  h += "</div>";
  box.innerHTML = h;
  cfgFnChanged();
}

// "count" counts rows, so it needs no column; "none" plots the column as it
// already stands.
function cfgFnChanged(){
  var w = $id("aggValWrap");
  if (!w) return;
  var fn = (($id("aggFn") || {}).value) || "count";
  w.style.display = (fn === "count") ? "none" : "";
}

function colOptions(cols, sel){
  var h = "";
  for (var i = 0; i < (cols || []).length; i++){
    h += "<option value='" + esc(cols[i]) + "'"
      + (cols[i] === sel ? " selected" : "") + ">" + esc(cols[i]) + "</option>";
  }
  return h || "<option value=''>(no columns known)</option>";
}

function typeOptions(list, sel){
  var h = "";
  for (var i = 0; i < list.length; i++){
    h += "<option value='" + esc(list[i]) + "'"
      + (list[i] === sel ? " selected" : "") + ">" + esc(list[i]) + "</option>";
  }
  return h;
}

// A catalog report declares its own filters. They were being applied from the
// definition's defaults and never offered, so a tile could only ever show the
// whole report.
// A report can lean on values that are set per church: what counts as
// serving, how long a background check lasts, when the fiscal year starts.
// Those have defaults, and a default is a guess. Say which ones this report
// uses and what they are set to, so a wrong number is traceable to a setting
// rather than assumed to be the data.
function settingsBanner(list){
  if (!list || !list.length) return "";
  var anyDefault = false;
  for (var i = 0; i < list.length; i++){ if (list[i].is_default) anyDefault = true; }
  var h = "<div style='margin-bottom:10px;padding:8px 10px;border:1px solid "
        + (anyDefault ? "#f0d9a0;background:#fdf7e8" : "#cfe2ff;background:#f2f7ff")
        + ";border-radius:4px;font-size:12px;'>"
        + "<b>This report uses church settings.</b> Check they are right before "
        + "trusting the numbers:<ul style='margin:6px 0 0 18px;padding:0;'>";
  for (var j = 0; j < list.length; j++){
    h += "<li>" + esc(list[j].label) + ": <b>" + esc(list[j].value) + "</b>"
      + (list[j].is_default
         ? " <span style='color:#8a5a00;'>(default, never set here)</span>" : "")
      + "</li>";
  }
  h += "</ul><div style='margin-top:6px;'>Change them under "
    + "<b>Setup</b>.</div></div>";
  return h;
}

function drawCatalogFilters(){
  var host = $id("cfgParams");
  if (!host) return;
  var ps = CFG.params || [];
  if (!ps.length){
    host.innerHTML = "<span class='db-muted'>This report has no filters.</span>";
    return;
  }
  var cur = CFG.filters || {};
  var h = "<div style='border-top:1px solid #e5e5e5;padding-top:10px;'>"
        + "<div style='font-weight:600;margin-bottom:6px;'>Filters</div>";
  for (var i = 0; i < ps.length; i++){
    var pp = ps[i], id = "cf_" + pp.name;
    var val = (cur[pp.name] !== undefined) ? cur[pp.name] : (pp["default"] || "");
    h += "<div style='margin-bottom:8px;'><label style='font-weight:normal;'>"
      + esc(pp.label || pp.name) + "</label>";
    if (pp.type === "daterange"){
      h += "<select id='" + id + "' class='cfp' data-n='" + esc(pp.name)
        + "' style='width:100%;padding:5px;'>"
        + dateRangeOptions(val) + "</select>"
        + "<div class='db-muted' style='font-size:11px;'>Ranges are worked out "
        + "when the tile runs, so a fiscal-year tile stays on the current year "
        + "without being edited.</div>";
    } else if (pp.type === "dropdown" || pp.type === "multi_select"){
      // Options come from the report's own source_sql, run on demand: the
      // values are church-defined and cannot travel in the catalog.
      h += "<select id='" + id + "' class='cfp' data-n='" + esc(pp.name) + "'"
        + (pp.type === "multi_select" ? " multiple size='4'" : "")
        + " style='width:100%;padding:5px;'>"
        + "<option value=''>-- all --</option></select>";
    } else {
      h += "<input id='" + id + "' class='cfp' data-n='" + esc(pp.name)
        + "' value='" + esc(val) + "' style='width:100%;padding:5px;'>";
    }
    h += "</div>";
  }
  h += "</div>";
  host.innerHTML = h;
  loadFilterOptions(ps, cur);
}

// Exactly the presets build_filter_sql implements. Offering any other value
// would look like a filter and quietly apply nothing.
var DATE_PRESETS = [["", "-- report default --"],
                    ["last_7_days", "last 7 days"],
                    ["last_30_days", "last 30 days"],
                    ["last_60_days", "last 60 days"],
                    ["last_90_days", "last 90 days"],
                    ["last_180_days", "last 180 days"],
                    ["last_365_days", "last 365 days"],
                    ["last_3_months", "last 3 months"],
                    ["last_6_months", "last 6 months"],
                    ["last_12_months", "last 12 months"],
                    ["last_24_months", "last 2 years"],
                    ["last_36_months", "last 3 years"],
                    ["last_48_months", "last 4 years"],
                    ["last_24_months", "last 2 years"],
                    ["last_36_months", "last 3 years"],
                    ["ytd", "calendar year to date"],
                    ["this_calendar_year", "this calendar year"],
                    ["last_calendar_year", "last calendar year"],
                    ["fiscal_ytd", "fiscal year to date"],
                    ["this_fiscal_year", "this fiscal year"],
                    ["last_fiscal_year", "last fiscal year"]];

function dateRangeOptions(sel){
  var h = "";
  for (var i = 0; i < DATE_PRESETS.length; i++){
    h += "<option value='" + DATE_PRESETS[i][0] + "'"
      + (DATE_PRESETS[i][0] === sel ? " selected" : "") + ">"
      + esc(DATE_PRESETS[i][1]) + "</option>";
  }
  return h;
}

function loadFilterOptions(ps, cur){
  var want = [];
  for (var i = 0; i < ps.length; i++){
    if (ps[i].type === "dropdown" || ps[i].type === "multi_select"){
      want.push(ps[i].name);
    }
  }
  if (!want.length) return;
  post({action: "filter_options", report_id: CFG.rid,
        names: want.join(",")}, function(r){
    if (!r || !r.success) return;
    for (var k = 0; k < want.length; k++){
      var el = $id("cf_" + want[k]);
      var list = (r.options || {})[want[k]] || [];
      if (!el) continue;
      var h = "<option value=''>-- all --</option>";
      var picked = String(cur[want[k]] === undefined ? "" : cur[want[k]]).split(",");
      for (var j = 0; j < list.length; j++){
        var isSel = false;
        for (var m = 0; m < picked.length; m++){
          if (picked[m] === String(list[j].value)) isSel = true;
        }
        h += "<option value='" + esc(list[j].value) + "'"
          + (isSel ? " selected" : "") + ">" + esc(list[j].label) + "</option>";
      }
      el.innerHTML = h;
    }
  });
}

function harvestCatalogFilters(){
  var out = {}, els = document.querySelectorAll("#cfgParams .cfp");
  for (var i = 0; i < els.length; i++){
    var el = els[i], nm = el.getAttribute("data-n");
    if (el.multiple){
      var picked = [];
      for (var o = 0; o < el.options.length; o++){
        if (el.options[o].selected && el.options[o].value) picked.push(el.options[o].value);
      }
      if (picked.length) out[nm] = picked.join(",");
    } else if (el.value !== undefined && String(el.value).length){
      out[nm] = el.value;
    }
  }
  return out;
}

var CFG = {rid: "", name: "", scopeable: false};

// Set while editing an existing tile, so the dialog saves back to it instead
// of appending a new one. -1 means "adding".
var EDIT_TILE = -1;

function dbEditTile(i){
  var t = tiles()[i];
  if (!t) return;
  if (t.kind === "links"){
    dbModal("Links", linksEditorHtml(t));
    LINK_TILE = i;
    lnkDraw();
    return;
  }
  if (t.kind !== "report"){
    // Widgets and built tiles carry their own definitions; only the title is
    // ours to change here.
    var nm = prompt("Tile title", t.title || "");
    if (nm === null) return;
    t.title = nm.trim() || t.title;
    markDirty();
    renderGrid();
    return;
  }
  EDIT_TILE = i;
  dbPickOpen(t.report_id, t.title || t.report_id, t);
}

function dbPickOpen(rid, name, existing){
  // The list page already asked how to show it. Rebuilding the select without
  // reading that threw the answer away and silently reverted to table.
  var picked = existing ? (existing.display || "")
                        : (($id("disp_" + rid) || {}).value || "");
  if (!existing) EDIT_TILE = -1;
  CFG = {rid: rid, name: name, scopeable: false, usesServing: false,
         source: (existing && existing.source) || "reports", disp: picked,
         types: [], columns: [], chartCfg: {},
         filters: (existing && existing.filters) || {},
         agg: (existing && existing.agg) || null,
         query: (existing && existing.query) || "",
         serving: (existing && existing.serving_types) || ""};
  for (var i = 0; i < (REPORTS || []).length; i++){
    if (REPORTS[i].id === rid){
      CFG.scopeable = !!REPORTS[i].scopeable;
      CFG.usesServing = !!REPORTS[i].uses_serving;
      CFG.source = REPORTS[i].source || "reports";
      CFG.installed = (REPORTS[i].installed !== false);
      CFG.types = REPORTS[i].display_types || ["table"];
      CFG.columns = REPORTS[i].columns || [];
      CFG.chartCfg = REPORTS[i].chart_cfg || {};
      CFG.params = REPORTS[i].params || [];
      CFG.settings = REPORTS[i].settings || [];
      if (!CFG.disp){ CFG.disp = REPORTS[i].default_display || "table"; }
    }
  }
  var h = "<div style='display:flex;gap:8px;align-items:center;"
        + "margin-bottom:10px;'>"
        + "<button class='btn btn-xs btn-default' onclick='showTilePicker()'>"
        + "&larr; Back</button><b>" + esc(name) + "</b></div>";
  h += "<div style='display:grid;grid-template-columns:1fr 1fr;gap:12px;'>"
     + "<div><label>Tile title</label>"
     + "<input id='cfgTitle' style='width:100%;padding:5px;' value='"
     + esc(name) + "'></div>"
     + "<div><label>Show as</label><select id='cfgDisp' "
     + "style='width:100%;padding:5px;'></select></div></div>";

  if (CFG.scopeable){
    // The report honors a people scope, so a saved search is the natural
    // control and its own filters are secondary.
    var searches = (CUSTOM_META && CUSTOM_META.searches) || [];
    h += "<div style='margin-top:12px;'><label>Limit to a Search Builder search</label>"
      + "<select id='cfgScope' style='width:100%;padding:5px;'>"
      + "<option value=''>-- everyone the report covers --</option>";
    for (var j = 0; j < searches.length; j++){
      h += "<option value='" + esc(searches[j].name) + "'>"
        + esc(searches[j].name) + "</option>";
    }
    h += "</select><div class='db-muted'>This report accepts a people scope, so "
      + "the search decides who it runs over.</div></div>";
  }

  if (CFG.usesServing){
    h += "<div style='margin-top:12px;'><label>What counts as serving</label>"
      + "<div class='db-muted'>This report counts people by member type. The "
      + "ids are created per church, so pick what serving means here. Leave it "
      + "on the church default unless this tile needs something narrower.</div>"
      + "<div id='cfgServing' style='margin-top:6px;'>"
      + "<span class='db-muted'>Loading member types...</span></div></div>";
  }

  h += "<div style='margin-top:12px;'><label>Report filters</label>"
     + "<div class='db-muted'>Whatever is set here is stored with the tile and "
     + "applied every time it loads.</div>"
     + "<div id='cfgFilters' style='margin-top:6px;'>"
     + "<span class='db-muted'>Loading filters...</span></div></div>";
  h += "<div style='margin-top:14px;'>"
     + "<button class='btn btn-sm btn-primary' onclick='dbPickAdd()'>"
     + (existing ? "Save tile" : "Add tile") + "</button>"
     + " <span id='cfgMsg' class='db-muted'></span></div>";
  dbModal("Add: " + name, h);

  if (CFG.usesServing) loadTileServing();

  if (CFG.source === "catalog"){
    // Its parameters travel with the definition; there is no remote panel.
    var ds = $id("cfgDisp");
    var ty = CFG.types && CFG.types.length ? CFG.types : ["table"];
    var dh = "";
    for (var t2 = 0; t2 < ty.length; t2++){
      dh += "<option value='" + esc(ty[t2]) + "'"
         + (ty[t2] === CFG.disp ? " selected" : "") + ">" + esc(ty[t2])
         + "</option>";
    }
    // A report with no chart configuration can still be charted, but only if
    // it is told what to group by and what to measure. Offering it without
    // that is what produced an axis full of PeopleIds.
    if (!contains(ty, "chart")){
      dh += "<option value='chart'" + (CFG.disp === "chart" ? " selected" : "")
         + ">chart (needs setup)</option>";
    }
    if (!contains(ty, "kpi")){
      dh += "<option value='kpi'" + (CFG.disp === "kpi" ? " selected" : "")
         + ">kpi (needs setup)</option>";
    }
    ds.innerHTML = dh;
    ds.setAttribute("onchange", "cfgDispChanged()");
    $id("cfgFilters").innerHTML = settingsBanner(CFG.settings)
      + "<div id='cfgAgg'></div>"
      + "<div id='cfgParams' style='margin-top:10px;'></div>";
    cfgDispChanged();
    drawCatalogFilters();
    return;
  }

  postReports({action: "load_filters", report_id: rid}, function(r){
    var box = $id("cfgFilters");
    if (!r || !r.success){
      box.innerHTML = "<span class='db-muted'>No filters available for this "
                    + "report.</span>";
    } else {
      box.innerHTML = r.filter_html || "<span class='db-muted'>This report has "
                    + "no filters.</span>";
      var ds = $id("cfgDisp");
      var types = r.display_types || ["table"];
      var dh = "";
      for (var k = 0; k < types.length; k++){
        dh += "<option value='" + esc(types[k]) + "'"
           + (types[k] === r.default_display ? " selected" : "") + ">"
           + esc(types[k]) + "</option>";
      }
      ds.innerHTML = dh;
      // Trust the report over the catalog: load_filters is authoritative.
      CFG.scopeable = !!r.bt_supported;
    }
  });
}

function loadTileServing(){
  var box = $id("cfgServing");
  if (!box) return;
  post({action: "member_types"}, function(mt){
    if (!mt || !mt.success){
      box.innerHTML = "<span class='db-err'>Could not read member types.</span>";
      return;
    }
    MTYPES = mt.types || [];
    postReports({action: "get_settings"}, function(r){
      var cur = "140,310,320,710";
      if (r && r.success && r.settings && r.settings.serving_member_types){
        cur = String(r.settings.serving_member_types);
      }
      drawTileServing(cur);
    });
  });
}

function drawTileServing(cur){
  var box = $id("cfgServing");
  if (!box) return;
  var on = {};
  var parts = String(cur).split(",");
  for (var i = 0; i < parts.length; i++){ on[parts[i].trim()] = true; }

  var byId = {}, bad = [], missing = [], missTotal = 0;
  for (var j = 0; j < MTYPES.length; j++){ byId[String(MTYPES[j].id)] = MTYPES[j]; }
  for (var k in on){
    if (!byId[k]) bad.push(k + " does not exist here");
    else if (!byId[k].people) bad.push(k + " (" + byId[k].name + ") is unused");
  }
  var rx = /leader|teacher|director|chairman|staff|volunteer|service|coach|host|helper/i;
  for (var m = 0; m < MTYPES.length; m++){
    var t = MTYPES[m];
    if (!on[String(t.id)] && t.people > 0 && rx.test(t.name)){
      missing.push(t); missTotal += t.people;
    }
  }

  var h = "";
  if (bad.length || missing.length){
    h += "<div class='db-err' style='border:1px solid #f3c0c4;background:#fdecea;"
      + "padding:8px;border-radius:4px;margin-bottom:8px;font-size:12px;'>"
      + "<b>This would undercount.</b> ";
    if (bad.length) h += esc(bad.join("; ")) + ". ";
    if (missing.length){
      var bits = [];
      for (var y = 0; y < missing.length; y++){
        bits.push(missing[y].name + " (" + missing[y].people + ")");
      }
      h += "Not counted: " + esc(bits.join(", ")) + ", about <b>"
        + missTotal + "</b> people.";
    }
    h += "</div>";
  }
  h += "<div style='max-height:180px;overflow:auto;border:1px solid #dde3ea;"
     + "border-radius:4px;padding:8px;display:grid;"
     + "grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:4px;'>";
  for (var a = 0; a < MTYPES.length; a++){
    var ty = MTYPES[a];
    h += "<label style='font-weight:normal;font-size:12px;'>"
      + "<input type='checkbox' class='cfgSv' value='" + ty.id + "'"
      + (on[String(ty.id)] ? " checked" : "") + " onclick='reDrawTileServing()'> "
      + esc(ty.name) + " <span class='db-muted'>(" + ty.people + ")</span></label>";
  }
  h += "</div>";
  box.innerHTML = h;
}

// Re-render so the warning tracks what is ticked right now.
function reDrawTileServing(){
  drawTileServing(pickedServing().join(","));
}

function pickedServing(){
  var out = [];
  var boxes = document.querySelectorAll(".cfgSv");
  for (var i = 0; i < boxes.length; i++){
    if (boxes[i].checked) out.push(boxes[i].value);
  }
  return out;
}

// The filter panel is markup Enterprise Reporting rendered, so read the values
// back off the controls rather than assuming a shape.
function harvestFilters(root){
  var out = {};
  if (!root) return out;
  var els = root.querySelectorAll("input[name], select[name], textarea[name]");
  for (var i = 0; i < els.length; i++){
    var e = els[i];
    var ty = (e.type || "").toLowerCase();
    if (ty === "button" || ty === "submit" || !e.name) continue;
    if (ty === "checkbox" || ty === "radio"){
      if (e.checked) out[e.name] = e.value;
    } else if (e.multiple && e.options){
      var vals = [];
      for (var j = 0; j < e.options.length; j++){
        if (e.options[j].selected) vals.push(e.options[j].value);
      }
      if (vals.length) out[e.name] = vals.join(",");
    } else if (e.value !== "" && e.value !== null && e.value !== undefined){
      out[e.name] = e.value;
    }
  }
  return out;
}

function dbPickAdd(){
  if (CFG.source === "catalog" && !CFG.installed){
    var msgEl = $id("cfgMsg");
    if (msgEl) msgEl.textContent = "Installing this report...";
    post({action: "catalog_install", ids: CFG.rid}, function(r){
      if (!r || !r.success || !(r.installed || []).length){
        alert((r && (r.error || (r.missing || []).join(", ")))
              || "Could not install this report.");
        if (msgEl) msgEl.textContent = "";
        return;
      }
      CFG.installed = true;
      // Mark it in place rather than dropping CATALOG, which would empty the
      // picker of every catalog report until the next page load.
      var cl = (CATALOG && CATALOG.reports) || [];
      for (var ci = 0; ci < cl.length; ci++){
        if (cl[ci].id === CFG.rid){ cl[ci].installed = true; }
      }
      dbPickAdd();
    });
    return;
  }
  var disp = (($id("cfgDisp") || {}).value) || "table";
  var title = (($id("cfgTitle") || {}).value || CFG.name).trim() || CFG.name;
  var scope = "";
  if (CFG.scopeable){
    scope = (($id("cfgScope") || {}).value) || "";
  }
  var filters = harvestFilters($id("cfgFilters"));
  var list = tiles();
  var maxY = 0;
  for (var i = 0; i < list.length; i++){
    maxY = Math.max(maxY, (list[i].y || 0) + (list[i].h || 4));
  }
  var serving = CFG.usesServing ? pickedServing().join(",") : "";
  var src = "reports";
  for (var q4 = 0; q4 < (REPORTS || []).length; q4++){
    if (REPORTS[q4].id === CFG.rid && REPORTS[q4].source){ src = REPORTS[q4].source; }
  }
  var agg = null;
  if (src === "catalog" && disp !== "table"){
    agg = {label: (($id("aggLabel") || {}).value) || "",
           fn: (($id("aggFn") || {}).value) || "count",
           value: (($id("aggVal") || {}).value) || "",
           chart: (($id("aggChart") || {}).value) || "bar",
           top: (($id("aggTop") || {}).value) || "15"};
  }
  // A catalog report's own filters are collected here; an Enterprise
  // Reporting one still uses its remote panel.
  if (src === "catalog"){ filters = harvestCatalogFilters(); }

  if (EDIT_TILE >= 0 && list[EDIT_TILE]){
    var t = list[EDIT_TILE];
    // Dropped BEFORE the edit: the key is derived from the tile's settings, so
    // after mutating it we would compute the new key and leave the stale entry
    // behind under the old one.
    cacheDrop(t);
    t.report_id = CFG.rid;
    t.source = src;
    t.title = title;
    t.display = disp;
    t.query = scope;
    t.filters = filters;
    t.serving_types = serving;
    t.agg = agg;
    EDIT_TILE = -1;
    dbCloseModal();
    markDirty();
    renderGrid();
    return;
  }

  list.push({kind: "report", report_id: CFG.rid, source: src,
             title: title + (scope ? (" - " + scope) : ""),
             display: disp, query: scope, filters: filters,
             serving_types: serving, agg: agg,
             x: 0, y: maxY, w: 4, h: 4,
             fit: (disp !== "chart" && disp !== "table")});
  dbCloseModal();
  markDirty();
  renderGrid();
}

function dbPick(rid, name){
  var sel = $id("disp_" + rid);
  var disp = sel ? sel.value : "table";
  // Only attach a scope the report will actually apply. Storing one on a
  // report that ignores people_ids would show a filter that does nothing.
  var scope = "";
  var scopeSel = $id("dbRepScope");
  if (scopeSel && scopeSel.value){
    for (var s2 = 0; s2 < (REPORTS || []).length; s2++){
      if (REPORTS[s2].id === rid && REPORTS[s2].scopeable){
        scope = scopeSel.value;
      }
    }
  }
  var list = tiles();
  // Drop it at the bottom so it never lands on top of an existing tile.
  var maxY = 0;
  for (var i = 0; i < list.length; i++){
    maxY = Math.max(maxY, (list[i].y || 0) + (list[i].h || 4));
  }
  list.push({kind: "report", report_id: rid,
             title: name + (scope ? (" - " + scope) : ""),
             display: disp, query: scope,
             x: 0, y: maxY, w: 4, h: 4, filters: {},
             fit: (disp !== "chart" && disp !== "table")});
  markDirty();
  dbCloseModal();
  renderGrid();
}

// ---------- modal ----------
function dbModal(title, html){
  $id("dbModalTitle").textContent = title;
  $id("dbModalBody").innerHTML = html;
  $id("dbModal").className = "db-modal on";
}
function dbCloseModal(){
  // Came from an expanded tile, so go back to it rather than dumping the user
  // on the dashboard with their selection gone.
  if (TASK_FROM_EXPAND){
    TASK_FROM_EXPAND = false;
    showExpand();
    return;
  }
  $id("dbModal").className = "db-modal";
  if (CHARTS["texp"]){
    try { CHARTS["texp"].destroy(); } catch (e) { }
    delete CHARTS["texp"];
  }
  EXP_IDX = -1;
}

// Somebody whose dashboards just disappeared should be told where they went
// on the screen they are already on, not have to find Settings.
function legacyNote(){
  var c = (DB_BOOT && DB_BOOT.legacy_count) || 0;
  var el = $id("dbLegacyNote");
  if (!el || !c) return;
  el.innerHTML = "<b>" + c + " dashboard(s) are waiting to be claimed.</b> "
    + "An earlier version could not tell users apart, so personal dashboards "
    + "were all saved to one place. They are hidden until someone says who "
    + "each belongs to. Nothing was deleted. "
    + (DB_BOOT.can_manage
       ? "<a href='#' onclick='dbSettings();return false;'>Sort them out</a>"
       : "If one of them is yours, an administrator can hand it back.");
  el.style.display = "block";
}

(function(){
  cachePurgeLegacy();
  legacyNote();
  var st = readUrlState();
  if (st.id){
    dbOpen(st.id, st.scope, st.tab);
  } else {
    dbHome();
  }
})();
</''' + '''script>
''')

    # Its own block: a failure inside the shared library must not take the
    # dashboard's script down with it.
    if update_js:
        page.append('<' + 'script>' + update_js + '</' + 'script>')

    model.Form = "".join(page)
