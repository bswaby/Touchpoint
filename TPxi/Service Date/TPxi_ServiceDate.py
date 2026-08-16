#roles=Admin
###############################################################################
# ServiceDate
###############################################################################
# Stores the most recent date a person Gave, went on a Mission Trip, or Served
# as a "ServiceDate" Extra Value (date type). Days-since is calculated on the
# fly in reports/queries:  DATEDIFF(day, ServiceDate, GETDATE())
#
# Written By: Ben Swaby (TPxi Software, LLC)
# Email: bswaby@gmail.com                                                                                                      
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
#==============================================================================
# INSTALLING
#==============================================================================
# 1. Admin > Advanced > Special Content > Python Scripts > New Python Script File
#    Name it TPxi_ServiceDays and paste this code in.
# 2. Open it once at /PyScriptForm/TPxi_ServiceDays and work through the
#    "Setup check" panel at the top of the page. It reads YOUR lookup tables and
#    tells you which ids to put in the CONFIGURATION block below. Do this before
#    the first run -- a wrong Serve type is silent, the script just never credits
#    anyone for serving.
# 3. Press "Preload" on that page to populate the Extra Value for everyone.
#    It runs in batches so it will not time out.
# 4. Add it to Morning Batch (see below) so it stays current.
#
#
# WHY THE DAILY RUN IS CHEAP
#   Morning batch does not fire at a predictable time -- it drifts. So the script 
#   does not assume "yesterday".
#   It records the timestamp of every successful run and, next time, only looks
#   at activity since then (minus an overlap, see INCREMENTAL_OVERLAP_DAYS).
#   Nothing is missed when a run is late or skipped for a week, and the daily
#   read drops from the whole lookback window to just what changed.
#
#   It automatically promotes itself to a full rescan when the watermark is
#   missing or older than FULL_RESCAN_EVERY_DAYS. That full pass is also the
#   only thing that can CLEAR a stale value, because a person aging out of the
#   window produces no new activity for an incremental pass to notice.
#
# MODES
#   /PyScriptForm/TPxi_ServiceDays          dashboard, setup check, and the
#                                           Preload / Force Full Rescan buttons
#   Data.run_servicedays = "true"           morning batch (incremental)
#   Data.run = "full" as well               makes that batch run a FULL rescan
#                                           instead, for a one-off from
#                                           the morning batch
#
# To rebuild everything by hand, use the Force Full Rescan button on the page.
# It batches the writes, so it will not time out the way a single big pass would.
#
###############################################################################
import json
import datetime


# =============================================================================
# CONFIGURATION  --  everything a different church needs to change lives here
# =============================================================================
# Open the dashboard once; the Setup check panel prints YOUR ids for each of
# these so you are not guessing.

# Every value below is a DEFAULT. Whatever you save from the Settings modal on
# the dashboard is stored in Special Content and overrides these at run time, so
# a church can configure the script without editing code. Editing here still
# works and is what a fresh install starts from.

CONFIG_DEFAULTS = {

    # --- What counts as "serving" --------------------------------------------
    # Ids from lookup.OrganizationType for your serving/volunteer involvements.
    # THIS IS THE ONE THAT MOST OFTEN NEEDS CHANGING and the one that fails
    # silently: if it is wrong the script runs happily and credits nobody for
    # serving. The Settings modal lists your types so you can tick them.
    'SERVE_ORG_TYPE_IDS': [145],           # FBCH: 145 = "Serve"

    # Mission trips use the built-in Organizations.IsMissionTrip flag, which is
    # standard across installs. False stops mission trips counting at all.
    'INCLUDE_MISSION_TRIPS': True,

    # --- What counts as "giving" ---------------------------------------------
    # Contribution types to EXCLUDE. Defaults are the non-gift types:
    #   6 Returned Check, 7 Reversed, 8 Pledge, 99 Event Registration Fee.
    # Verify against your own lookup.ContributionType -- these ids do vary.
    'EXCLUDE_CONTRIBUTION_TYPE_IDS': [6, 7, 8, 99],

    # Only count posted contributions. 0 = posted in a standard install.
    'CONTRIBUTION_STATUS_ID': 0,

    # Give a spouse credit for the household's giving. Most churches want this
    # on, because giving is usually recorded against one member of a couple.
    'CREDIT_SPOUSE_FOR_GIVING': True,

    # --- Scope and naming -----------------------------------------------------
    'SCOPE_DAYS': 730,                     # how far back a FULL rescan looks
    'EV_NAME': "ServiceDate",              # the Extra Value this maintains
    'OLD_EV_NAME': "ServiceDays",          # legacy int EV, offered for cleanup

    # --- Search Builder "In SQL List" scripts ---------------------------------
    # This script only maintains the Extra Value. What staff actually USE is
    # Search Builder, and to get there they need a saved SQL script per window.
    # These are the day windows to publish, one script each, named EV_NAME+days
    # (ServiceDate90, ServiceDate180, ServiceDate365). Add or remove numbers and
    # the page will offer to create the difference.
    'SQL_LIST_WINDOWS': [90, 180, 365],

    # --- Incremental behaviour ------------------------------------------------
    # Overlap re-examined on every incremental run. Covers contributions entered
    # today but dated earlier, and attendance posted late. Do not set to 0.
    'INCREMENTAL_OVERLAP_DAYS': 2,

    # Force a full rescan when the last one is older than this. The full pass is
    # what clears people who aged out of SCOPE_DAYS, which an incremental pass
    # cannot see. 7 = a full pass roughly weekly.
    'FULL_RESCAN_EVERY_DAYS': 7,

    # Name of the Data flag your morning batch sets.
    'BATCH_TRIGGER': "run_servicedays",

    # Which Special Content Python script runs this daily. On a standard install
    # that is "MorningBatch" -- the one that fires every morning. It is NOT
    # "ScheduledTasks", which is a separate feature for jobs you schedule
    # yourself; putting the trigger there looks installed but never runs daily.
    'BATCH_CONTENT_NAME': "MorningBatch",
}

# Special Content records. Created automatically on first save.
CONFIG_CONTENT_NAME = "TPxi_ServiceDays_Config"
STATE_CONTENT_NAME  = "TPxi_ServiceDays_State"


def _coerce(key, value, default):
    """Settings arrive from an HTML form as strings. Coerce to the shape of the
    default so the SQL builders can keep assuming real types."""
    try:
        if isinstance(default, bool):
            if isinstance(value, bool):
                return value
            return str(value).strip().lower() in ('1', 'true', 'yes', 'on')
        if isinstance(default, list):
            if isinstance(value, list):
                raw = value
            else:
                raw = [x for x in str(value).replace(';', ',').split(',')]
            out = []
            for x in raw:
                x = str(x).strip()
                if x == '':
                    continue
                try:
                    out.append(int(x))
                except:
                    # Skip the one bad entry rather than throwing the whole
                    # list away. Discarding everything would silently revert
                    # the field to its default and the user would believe the
                    # ids they typed had been saved.
                    continue
            return out
        if isinstance(default, int):
            return int(str(value).strip())
        return str(value)
    except:
        return default


def load_config():
    """Saved settings merged over CONFIG_DEFAULTS."""
    cfg = {}
    for k, v in CONFIG_DEFAULTS.items():
        cfg[k] = v
    try:
        raw = model.TextContent(CONFIG_CONTENT_NAME)
        if raw:
            saved = json.loads(raw)
            for k in CONFIG_DEFAULTS:
                if k in saved:
                    cfg[k] = _coerce(k, saved[k], CONFIG_DEFAULTS[k])
    except:
        # A corrupt or missing config must never stop the script -- fall back to
        # defaults rather than failing the morning batch.
        pass
    return cfg


def save_config(new_values):
    """Persist only the recognised keys, coerced to the right types."""
    cfg = load_config()
    for k in CONFIG_DEFAULTS:
        if k in new_values:
            cfg[k] = _coerce(k, new_values[k], CONFIG_DEFAULTS[k])
    model.WriteContentText(CONFIG_CONTENT_NAME, json.dumps(cfg), "")
    return cfg


# Bind the merged config to module-level names so every existing reference in
# the SQL builders keeps working untouched.
_CFG = load_config()
SERVE_ORG_TYPE_IDS            = _CFG['SERVE_ORG_TYPE_IDS']
INCLUDE_MISSION_TRIPS         = _CFG['INCLUDE_MISSION_TRIPS']
EXCLUDE_CONTRIBUTION_TYPE_IDS = _CFG['EXCLUDE_CONTRIBUTION_TYPE_IDS']
CONTRIBUTION_STATUS_ID        = _CFG['CONTRIBUTION_STATUS_ID']
CREDIT_SPOUSE_FOR_GIVING      = _CFG['CREDIT_SPOUSE_FOR_GIVING']
SCOPE_DAYS                    = _CFG['SCOPE_DAYS']
EV_NAME                       = _CFG['EV_NAME']
OLD_EV_NAME                   = _CFG['OLD_EV_NAME']
INCREMENTAL_OVERLAP_DAYS      = _CFG['INCREMENTAL_OVERLAP_DAYS']
FULL_RESCAN_EVERY_DAYS        = _CFG['FULL_RESCAN_EVERY_DAYS']
BATCH_TRIGGER                 = _CFG['BATCH_TRIGGER']
SQL_LIST_WINDOWS              = _CFG['SQL_LIST_WINDOWS']
BATCH_CONTENT_NAME            = _CFG['BATCH_CONTENT_NAME']


# =============================================================================
# MORNING BATCH REGISTRATION
# =============================================================================
# Same managed-block pattern the other TPxi scripts use: our lines live between
# two markers so they can be added and removed without disturbing anything else
# sharing the slot.

_SCHED_MARKER_START = "# >>> TPxi_ServiceDays schedule start (managed by app, do not edit) >>>"
_SCHED_MARKER_END   = "# <<< TPxi_ServiceDays schedule end <<<"
_SCHED_CONTENT_SLOT = BATCH_CONTENT_NAME
_DEFAULT_SCRIPT_NAME = "TPxi_ServiceDays"

# Slots that are NOT the configured one but are worth checking anyway, because
# a block left in one of them is the failure you cannot see: the page can say
# "installed" while nothing runs daily. "ScheduledTasks" is the trap - it is a
# separate TouchPoint feature, not the morning batch.
_SCHED_OTHER_SLOTS = [s for s in ("MorningBatch", "ScheduledTasks")
                      if s != _SCHED_CONTENT_SLOT]


def get_script_name():
    """What this script is actually called on THIS install.

    The name matters because ScheduledTasks has to CallScript it by name, and
    admins rename scripts.

    TouchPoint does NOT reliably expose the running script's name to Python --
    model.URL is often empty here -- so on a plain page load this falls through
    to the default. That is why the page detects the real name in JavaScript
    (window.location.pathname) and posts it as script_name on every AJAX call,
    which is what the install action actually registers. The displayed name is
    corrected client-side for the same reason.

    Resolution order:
      1. script_name posted by the page (authoritative)
      2. model.URL, when TouchPoint happens to provide it
      3. the default
    """
    try:
        if hasattr(Data, 'script_name'):
            sn = str(getattr(Data, 'script_name', '') or '').strip()
            if sn:
                return sn
    except:
        pass
    try:
        import re as _re
        url = str(getattr(model, 'URL', '') or '')
        m = _re.search(r'/PyScript(?:Form)?/([^/?#&]+)', url)
        if m:
            return m.group(1)
    except:
        pass
    return _DEFAULT_SCRIPT_NAME


def _sched_block(script_name):
    return (_SCHED_MARKER_START + "\n"
            "try:\n"
            "    Data." + BATCH_TRIGGER + " = 'true'\n"
            "    model.CallScript('" + script_name + "')\n"
            "except Exception as _sd_e:\n"
            "    print 'ServiceDate batch error: ' + str(_sd_e)\n"
            + _SCHED_MARKER_END + "\n")


def _sched_read(slot):
    try:
        return model.PythonContent(slot) or "", True
    except:
        return "", False


def _sched_registered_name(text):
    """The script name actually inside our managed block.

    This is the difference between what we GUESS we are called and what is
    really registered to run. On a plain page load get_script_name() falls back
    to the default, so reporting that as the installed name shows the default
    forever no matter what was installed.
    """
    try:
        import re as _re
        pat = (_re.escape(_SCHED_MARKER_START) + r".*?CallScript\(\s*['\"]([^'\"]+)['\"]"
               r".*?" + _re.escape(_SCHED_MARKER_END))
        m = _re.search(pat, text, _re.DOTALL)
        if m:
            return m.group(1)
    except:
        pass
    return ''


def sched_status():
    """Is our block installed, where, and under what name?"""
    sn = get_script_name()
    existing, ok = _sched_read(_SCHED_CONTENT_SLOT)
    if not ok:
        return {'readable': False, 'installed': False, 'outside': False,
                'script_name': sn, 'registered_name': '',
                'slot': _SCHED_CONTENT_SLOT, 'stray': []}

    installed = _SCHED_MARKER_START in existing
    registered = _sched_registered_name(existing) if installed else ''

    outside = False
    try:
        stripped = existing
        if installed:
            import re as _re
            pat = _re.escape(_SCHED_MARKER_START) + r".*?" + _re.escape(_SCHED_MARKER_END)
            stripped = _re.sub(pat, "", existing, flags=_re.DOTALL)
        # A hand-rolled entry someone added earlier. Worth flagging, because
        # installing ours on top would make it run twice.
        if ("CallScript('" + sn + "')") in stripped or \
           ('CallScript("' + sn + '")') in stripped:
            outside = True
    except:
        pass

    # A block sitting in the wrong slot reports "installed" but never runs
    # daily, so surface it rather than leaving it to be discovered by the
    # ServiceDate quietly going stale.
    stray = []
    for other in _SCHED_OTHER_SLOTS:
        text, ok2 = _sched_read(other)
        if ok2 and _SCHED_MARKER_START in text:
            stray.append(other)

    return {'readable': True, 'installed': installed, 'outside': outside,
            'script_name': sn, 'registered_name': registered,
            'slot': _SCHED_CONTENT_SLOT, 'stray': stray}


def sched_install():
    sn = get_script_name()
    try:
        existing = model.PythonContent(_SCHED_CONTENT_SLOT) or ""
    except Exception as e:
        return {'success': False, 'message': 'Could not read ' + _SCHED_CONTENT_SLOT + ': ' + str(e)}
    if _SCHED_MARKER_START in existing:
        return {'success': True, 'message': 'Already installed.', 'already': True}
    new_content = existing.rstrip() + ("\n\n" if existing.strip() else "") + _sched_block(sn)
    try:
        model.WriteContentPython(_SCHED_CONTENT_SLOT, new_content)
        return {'success': True,
                'message': 'Added to ' + _SCHED_CONTENT_SLOT + '. It will run on the next morning batch.'}
    except Exception as e:
        return {'success': False, 'message': 'Could not write ' + _SCHED_CONTENT_SLOT + ': ' + str(e)}


def _sched_strip(slot):
    """Remove our managed block from one slot. Returns (changed, error)."""
    existing, ok = _sched_read(slot)
    if not ok:
        return False, 'Could not read ' + slot
    if _SCHED_MARKER_START not in existing:
        return False, ''
    try:
        import re as _re
        pat = _re.escape(_SCHED_MARKER_START) + r".*?" + _re.escape(_SCHED_MARKER_END) + r"\n?"
        model.WriteContentPython(
            slot, _re.sub(pat, "", existing, flags=_re.DOTALL).rstrip() + "\n")
        return True, ''
    except Exception as e:
        return False, 'Could not write ' + slot + ': ' + str(e)


def sched_uninstall():
    """Remove from the configured slot AND any other slot holding our block.

    Sweeping the others matters because an earlier version of this page wrote
    to ScheduledTasks. Removing only the configured slot would leave that one
    behind, invisible, forever.
    """
    removed, errors = [], []
    for slot in [_SCHED_CONTENT_SLOT] + _SCHED_OTHER_SLOTS:
        changed, err = _sched_strip(slot)
        if changed:
            removed.append(slot)
        if err:
            errors.append(err)
    if errors and not removed:
        return {'success': False, 'message': '; '.join(errors)}
    if not removed:
        return {'success': True, 'message': 'Nothing to remove.'}
    msg = 'Removed from ' + ', '.join(removed) + '.'
    if errors:
        msg += ' (' + '; '.join(errors) + ')'
    return {'success': len(errors) == 0, 'message': msg}


# =============================================================================
# SEARCH BUILDER "IN SQL LIST" SCRIPTS
# =============================================================================
# The Extra Value this script maintains is invisible to most staff. What they
# actually use is Search Builder > In SQL List, which needs a saved SQL script
# per time window. Those are ordinary Special Content SQL records (TypeID 4),
# so we can check for them and create them.
#
# Deliberately NOT auto-created on page load: writing Special Content behind
# someone's back is the kind of thing you want to have clicked a button for.

def _sqllist_name(days):
    """ServiceDate90, ServiceDate180, ... Derived from EV_NAME so a church that
    renames the Extra Value gets matching script names."""
    return "{0}{1}".format(EV_NAME, int(days))


def _sqllist_body(days):
    """The query behind one window.

    DateValue >= today - N. Deceased and archived people are excluded here as
    well as in the maintenance pass, because the Extra Value lingers on someone
    who dies until the next full rescan clears it.
    """
    return (
        "SELECT DISTINCT p.PeopleId\n"
        "FROM People p\n"
        "JOIN PeopleExtra pe ON pe.PeopleId = p.PeopleId\n"
        "WHERE pe.Field = '{0}'\n"
        "AND pe.DateValue >= DATEADD(day, -{1}, GETDATE())\n"
        "AND p.IsDeceased = 0\n"
        "AND p.ArchivedFlag = 0"
    ).format(EV_NAME.replace("'", "''"), int(days))


def _sqllist_norm(s):
    """Compare bodies ignoring line-ending and trailing-whitespace noise, so a
    script saved through the TouchPoint editor still reads as unchanged."""
    try:
        lines = str(s or "").replace('\r\n', '\n').replace('\r', '\n').split('\n')
        return '\n'.join([ln.rstrip() for ln in lines]).strip()
    except:
        return ""


def sqllist_status():
    """For each window: does the script exist, and does it still match ours?

    Three states per script, and the middle one is the point of this function:
      missing  - not there, we can create it
      ok       - present and identical to what we would write
      differs  - present but edited (by them, or by an older EV_NAME). We do
                 NOT quietly overwrite that; the page offers it as a choice.
    """
    items = []
    readable = True
    for days in SQL_LIST_WINDOWS:
        name = _sqllist_name(days)
        try:
            body = model.SqlContent(name) or ""
        except:
            readable = False
            body = ""
        if not str(body).strip():
            state = 'missing'
        elif _sqllist_norm(body) == _sqllist_norm(_sqllist_body(days)):
            state = 'ok'
        else:
            state = 'differs'
        items.append({'days': int(days), 'name': name, 'state': state})
    return {
        'readable': readable,
        'items': items,
        'missing': len([i for i in items if i['state'] == 'missing']),
        'differs': len([i for i in items if i['state'] == 'differs']),
        'ok': len([i for i in items if i['state'] == 'ok']),
    }


def sqllist_install(overwrite=False):
    """Create the missing scripts. Only touches edited ones when overwrite."""
    created, replaced, skipped, failed = [], [], [], []
    for item in sqllist_status()['items']:
        days, name, state = item['days'], item['name'], item['state']
        if state == 'ok':
            continue
        if state == 'differs' and not overwrite:
            skipped.append(name)
            continue
        try:
            model.WriteContentSql(name, _sqllist_body(days), "ServiceDate")
            (replaced if state == 'differs' else created).append(name)
        except Exception as e:
            failed.append("{0}: {1}".format(name, str(e)))

    parts = []
    if created:  parts.append("Created " + ", ".join(created) + ".")
    if replaced: parts.append("Replaced " + ", ".join(replaced) + ".")
    if skipped:  parts.append("Left " + ", ".join(skipped) +
                              " alone (edited here - use Overwrite to replace).")
    if failed:   parts.append("Failed - " + "; ".join(failed))
    if not parts:
        parts.append("Nothing to do; all scripts are already correct.")
    return {'success': len(failed) == 0, 'message': " ".join(parts),
            'status': sqllist_status()}


# =============================================================================
# RUN STATE  --  the last-run watermark
# =============================================================================
# Persisted because morning batch timing is unreliable. Storing the actual
# completion time means a late, skipped, or failed run costs nothing: the next
# run simply reaches further back.

def load_state():
    try:
        raw = model.TextContent(STATE_CONTENT_NAME)
        if raw:
            return json.loads(raw)
    except:
        pass
    return {}


def save_state(state):
    try:
        model.WriteContentText(STATE_CONTENT_NAME, json.dumps(state), "")
    except Exception as e:
        # Never let a bookkeeping failure lose the work that was just done.
        # Worst case the next run does a full pass.
        print "<!-- ServiceDate: could not save state: {0} -->".format(str(e))


def get_last_run():
    """datetime of the last successful run, or None."""
    raw = (load_state() or {}).get('last_run')
    if not raw:
        return None
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d'):
        try:
            return datetime.datetime.strptime(str(raw)[:19], fmt)
        except:
            continue
    return None


def record_run(mode, updated, cleared, scanned_from):
    state = load_state() or {}
    state['last_run'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    state['last_mode'] = mode
    state['last_updated'] = updated
    state['last_cleared'] = cleared
    state['last_scanned_from'] = scanned_from or ''
    if mode == 'full':
        state['last_full_run'] = state['last_run']
    save_state(state)


def decide_mode(force_full=False):
    """Returns (mode, since_datetime_or_None, why).

    Incremental is only safe when we know when we last ran AND a full pass has
    happened recently enough to have cleared anything that aged out.
    """
    if force_full:
        return ('full', None, 'forced from the dashboard')
    last = get_last_run()
    if last is None:
        return ('full', None, 'no previous run recorded')
    state = load_state() or {}
    last_full_raw = state.get('last_full_run')
    last_full = None
    if last_full_raw:
        try:
            last_full = datetime.datetime.strptime(str(last_full_raw)[:19], '%Y-%m-%d %H:%M:%S')
        except:
            last_full = None
    if last_full is None:
        return ('full', None, 'no full rescan on record')
    age = (datetime.datetime.now() - last_full).days
    if age >= FULL_RESCAN_EVERY_DAYS:
        return ('full', None, 'last full rescan was {0} days ago'.format(age))
    since = last - datetime.timedelta(days=INCREMENTAL_OVERLAP_DAYS)
    return ('incremental', since, 'activity since {0}'.format(since.strftime('%Y-%m-%d %H:%M')))

# ---------------------------------------------------------------------------
# Core SQL
# ---------------------------------------------------------------------------
# Two shapes from one builder:
#   FULL         - every person, every activity inside SCOPE_DAYS. Can also
#                  CLEAR people who have aged out.
#   INCREMENTAL  - only people with activity since the watermark. Much cheaper.
#                  Never clears, and never moves a date backwards (see below).
#
# Why incremental only moves dates FORWARD: a contribution entered today but
# dated three months ago is real and must be picked up, but it must not
# overwrite a more recent ServiceDate the person already has. So the UPDATE
# test is "> existing", not "!= existing" as the full pass uses.

def _csv(vals, fallback='-1'):
    """Render a config list into a SQL IN() list, never empty."""
    out = [str(int(v)) for v in (vals or [])]
    return ','.join(out) if out else fallback


def _serve_clause():
    """Serving involvements. Empty config disables the Serve source entirely
    rather than silently matching everything."""
    if not SERVE_ORG_TYPE_IDS:
        return "AND 1 = 0"
    return "AND o.OrganizationTypeId IN ({0})".format(_csv(SERVE_ORG_TYPE_IDS))


def _spouse_credit_sql():
    """The UNION arms that give a spouse credit for household giving."""
    if not CREDIT_SPOUSE_FOR_GIVING:
        return ""
    return """
        UNION ALL

        SELECT f.HeadOfHouseholdSpouseId as PeopleId, lg.LastGaveDate
        FROM LastGiving lg
        JOIN Families f ON lg.PeopleId = f.HeadOfHouseholdId
        WHERE f.HeadOfHouseholdSpouseId IS NOT NULL
            AND f.HeadOfHouseholdSpouseId > 0

        UNION ALL

        SELECT f.HeadOfHouseholdId as PeopleId, lg.LastGaveDate
        FROM LastGiving lg
        JOIN Families f ON lg.PeopleId = f.HeadOfHouseholdSpouseId
        WHERE f.HeadOfHouseholdId IS NOT NULL
            AND f.HeadOfHouseholdId > 0
    """


def get_sql(since=None):
    """since=None -> full rescan. since=datetime -> incremental."""
    incremental = since is not None

    if incremental:
        floor_sql = "'{0}'".format(since.strftime('%Y-%m-%d %H:%M:%S'))
        date_floor_contrib = "c.ContributionDate >= {0}".format(floor_sql)
        date_floor_meeting = "m.MeetingDate >= {0}".format(floor_sql)
    else:
        date_floor_contrib = "c.ContributionDate >= DATEADD(day, -{0}, GETDATE())".format(SCOPE_DAYS)
        date_floor_meeting = "m.MeetingDate >= DATEADD(day, -{0}, GETDATE())".format(SCOPE_DAYS)

    mission_clause = "AND o.IsMissionTrip = 1" if INCLUDE_MISSION_TRIPS else "AND 1 = 0"

    # Incremental restricts to people who actually have new activity. The full
    # pass walks People so it can also spot values that need clearing.
    if incremental:
        person_source = """
        AffectedPeople AS (
            SELECT PeopleId FROM Giving
            UNION SELECT PeopleId FROM Mission
            UNION SELECT PeopleId FROM Serve
        ),"""
        person_join = "JOIN AffectedPeople ap ON ap.PeopleId = p.PeopleId"
    else:
        person_source = ""
        person_join = ""

    # Full pass: any difference is a change, and a missing calc means CLEAR.
    # Incremental: forward-only, and never clear.
    if incremental:
        action_case = """
        CASE
            WHEN c.CalcServiceDate IS NOT NULL
                AND (ev.DateValue IS NULL OR c.CalcServiceDate > CAST(ev.DateValue AS DATE))
                THEN 'UPDATE'
            ELSE 'OK'
        END as Action"""
        final_where = "WHERE c.CalcServiceDate IS NOT NULL"
    else:
        action_case = """
        CASE
            WHEN c.CalcServiceDate IS NOT NULL
                AND (ev.DateValue IS NULL OR CAST(ev.DateValue AS DATE) != c.CalcServiceDate)
                THEN 'UPDATE'
            WHEN c.CalcServiceDate IS NULL AND ev.DateValue IS NOT NULL
                THEN 'CLEAR'
            ELSE 'OK'
        END as Action"""
        final_where = """WHERE c.CalcServiceDate IS NOT NULL
        OR ev.DateValue IS NOT NULL"""

    return """
    ;WITH Today AS (
        SELECT CAST(GETDATE() AS DATE) as dt
    ),
    LastGiving AS (
        SELECT
            c.PeopleId,
            MAX(CAST(c.ContributionDate AS DATE)) as LastGaveDate
        FROM Contribution c
        WHERE c.ContributionTypeId NOT IN ({exclude_types})
            AND c.ContributionStatusId = {status_id}
            AND {floor_contrib}
            AND c.ContributionAmount > 0
        GROUP BY c.PeopleId
    ),
    GivingWithSpouse AS (
        SELECT lg.PeopleId, lg.LastGaveDate
        FROM LastGiving lg
        {spouse_credit}
    ),
    Giving AS (
        SELECT PeopleId, MAX(LastGaveDate) as LastGaveDate
        FROM GivingWithSpouse
        GROUP BY PeopleId
    ),
    Mission AS (
        SELECT
            a.PeopleId,
            MAX(CAST(m.MeetingDate AS DATE)) as LastMissionDate
        FROM Attend a
        JOIN Meetings m ON a.MeetingId = m.MeetingId
        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
        WHERE 1 = 1
            {mission_clause}
            AND a.AttendanceFlag = 1
            AND {floor_meeting}
        GROUP BY a.PeopleId
    ),
    Serve AS (
        SELECT
            a.PeopleId,
            MAX(CAST(m.MeetingDate AS DATE)) as LastServeDate
        FROM Attend a
        JOIN Meetings m ON a.MeetingId = m.MeetingId
        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
        WHERE 1 = 1
            {serve_clause}
            AND a.AttendanceFlag = 1
            AND {floor_meeting}
        GROUP BY a.PeopleId
    ),{person_source}
    Calculated AS (
        SELECT
            p.PeopleId,
            p.Name2 as PersonName,
            ms.Description as MemberStatus,
            fp.Description as FamilyPosition,
            g.LastGaveDate,
            mi.LastMissionDate,
            s.LastServeDate,
            (
                SELECT MAX(v)
                FROM (VALUES
                    (g.LastGaveDate),
                    (mi.LastMissionDate),
                    (s.LastServeDate)
                ) AS T(v)
                WHERE v IS NOT NULL
            ) as CalcServiceDate,
            CASE
                WHEN g.LastGaveDate IS NOT NULL
                    AND g.LastGaveDate >= ISNULL(mi.LastMissionDate, '1900-01-01')
                    AND g.LastGaveDate >= ISNULL(s.LastServeDate, '1900-01-01')
                    THEN 'Gave'
                WHEN mi.LastMissionDate IS NOT NULL
                    AND mi.LastMissionDate >= ISNULL(g.LastGaveDate, '1900-01-01')
                    AND mi.LastMissionDate >= ISNULL(s.LastServeDate, '1900-01-01')
                    THEN 'Mission'
                WHEN s.LastServeDate IS NOT NULL
                    THEN 'Serve'
                ELSE NULL
            END as Driver,
            t.dt as TodayDate
        FROM People p
        CROSS JOIN Today t
        {person_join}
        LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
        LEFT JOIN lookup.FamilyPosition fp ON p.PositionInFamilyId = fp.Id
        LEFT JOIN Giving g ON p.PeopleId = g.PeopleId
        LEFT JOIN Mission mi ON p.PeopleId = mi.PeopleId
        LEFT JOIN Serve s ON p.PeopleId = s.PeopleId
        WHERE p.IsDeceased = 0
            AND p.ArchivedFlag = 0
    )
    SELECT
        c.PeopleId,
        c.PersonName,
        c.MemberStatus,
        c.FamilyPosition,
        c.LastGaveDate,
        c.LastMissionDate,
        c.LastServeDate,
        c.CalcServiceDate,
        c.Driver,
        c.TodayDate,
        CASE WHEN c.CalcServiceDate IS NOT NULL
             THEN CASE WHEN c.CalcServiceDate > c.TodayDate THEN 0
                       ELSE DATEDIFF(day, c.CalcServiceDate, c.TodayDate)
                  END
             ELSE NULL
        END as DaysSince,
        ev.DateValue as ExistingServiceDate,
        {action_case}
    FROM Calculated c
    LEFT JOIN PeopleExtra ev ON c.PeopleId = ev.PeopleId
        AND ev.Field = '{ev_name}'
    {final_where}
    ORDER BY
        CASE WHEN c.CalcServiceDate IS NOT NULL THEN 0 ELSE 1 END,
        c.CalcServiceDate DESC
    """.format(
        exclude_types=_csv(EXCLUDE_CONTRIBUTION_TYPE_IDS, '-1'),
        status_id=int(CONTRIBUTION_STATUS_ID),
        floor_contrib=date_floor_contrib,
        floor_meeting=date_floor_meeting,
        spouse_credit=_spouse_credit_sql(),
        mission_clause=mission_clause,
        serve_clause=_serve_clause(),
        person_source=person_source,
        person_join=person_join,
        action_case=action_case,
        ev_name=EV_NAME,
        final_where=final_where)


# ---------------------------------------------------------------------------
# SQL to find legacy ServiceDays int EVs to clean up
# ---------------------------------------------------------------------------
def get_cleanup_sql():
    return """
    SELECT pe.PeopleId, p.Name2 as PersonName, pe.IntValue
    FROM PeopleExtra pe
    JOIN People p ON pe.PeopleId = p.PeopleId
    WHERE pe.Field = '{0}'
        AND pe.IntValue IS NOT NULL
    ORDER BY p.Name2
    """.format(OLD_EV_NAME)


# ===========================================================================
# Determine mode
# ===========================================================================
is_batch = hasattr(Data, BATCH_TRIGGER) and str(getattr(Data, BATCH_TRIGGER)) == "true"
force_full = hasattr(Data, 'run') and str(Data.run) == 'full'
is_ajax = model.HttpMethod == "post" and hasattr(Data, 'action')

# ===========================================================================
# AJAX HANDLER
# ===========================================================================
if is_ajax:
    action = str(Data.action)

    if action == "preload":
        batch_size = 1000

        # Query only returns UPDATE/CLEAR rows - each call processes the next batch
        # As records get updated they become OK and drop out of the result set
        results = q.QuerySql(get_sql())

        needs_work = []
        for row in results:
            if row.Action == "UPDATE":
                needs_work.append(("UPDATE", row.PeopleId, row.CalcServiceDate))
            elif row.Action == "CLEAR":
                needs_work.append(("CLEAR", row.PeopleId, None))

        total = len(needs_work)
        batch = needs_work[:batch_size]

        updated = 0
        cleared = 0
        for action_type, pid, sdate in batch:
            if action_type == "UPDATE":
                model.AddExtraValueDate(pid, EV_NAME, sdate)
                updated += 1
            elif action_type == "CLEAR":
                model.DeleteExtraValue(pid, EV_NAME)
                cleared += 1

        remaining = total - len(batch)

        # A completed preload IS a full rescan, so stamp the watermark. Without
        # this the first morning batch after a preload would redo the whole
        # thing instead of going incremental.
        if remaining <= 0:
            record_run('full', updated, cleared, '')

        response = {
            "success": True,
            "updated": updated,
            "cleared": cleared,
            "processed": len(batch),
            "remaining": remaining,
            "total_found": total,
            "has_more": remaining > 0
        }
        print json.dumps(response)

    elif action == "save_config":
        # The modal posts one field per setting. Unknown keys are ignored and
        # everything is coerced to the type of its default.
        incoming = {}
        for k in CONFIG_DEFAULTS:
            if hasattr(Data, k):
                incoming[k] = str(getattr(Data, k))
        # Checkboxes only post when ticked, so an absent boolean means False.
        for k in CONFIG_DEFAULTS:
            if isinstance(CONFIG_DEFAULTS[k], bool) and k not in incoming:
                incoming[k] = False
        try:
            saved = save_config(incoming)
            print json.dumps({"success": True, "config": saved,
                              "message": "Settings saved. Run a Force Full Rescan "
                                         "so existing values pick up the change."})
        except Exception as e:
            print json.dumps({"success": False, "message": str(e)})

    elif action == "sched_install":
        print json.dumps(sched_install())

    elif action == "sched_uninstall":
        print json.dumps(sched_uninstall())

    elif action == "sched_status":
        print json.dumps(sched_status())

    elif action == "sqllist_install":
        ow = str(getattr(Data, 'overwrite', '') or '').strip().lower() in ('1', 'true', 'yes', 'on')
        print json.dumps(sqllist_install(ow))

    elif action == "sqllist_status":
        print json.dumps(sqllist_status())

    elif action == "cleanup":
        # Remove legacy ServiceDays (int) Extra Values
        batch_size = 1000

        results = q.QuerySql(get_cleanup_sql())
        all_ids = [row.PeopleId for row in results]

        total = len(all_ids)
        batch = all_ids[:batch_size]

        removed = 0
        for pid in batch:
            model.DeleteExtraValue(pid, OLD_EV_NAME)
            removed += 1

        remaining = total - len(batch)
        response = {
            "success": True,
            "removed": removed,
            "remaining": remaining,
            "total_found": total,
            "has_more": remaining > 0
        }
        print json.dumps(response)


# ===========================================================================
# MORNING BATCH MODE - Maintenance (deltas only)
# ===========================================================================
elif is_batch:
    # Incremental unless the watermark says otherwise. decide_mode() promotes
    # itself to a full rescan when there is no watermark, no full pass on
    # record, or the last full pass has aged out -- so a missed week of morning
    # batches repairs itself rather than leaving a hole.
    mode, since, why = decide_mode(force_full=force_full)

    updated = 0
    cleared = 0
    errors = 0
    queried = False

    # Belt and braces. The ScheduledTasks block wraps the CallScript, but this
    # guard means a failure here degrades to a logged line instead of an
    # exception crossing back into morning batch at all.
    try:
        results = q.QuerySql(get_sql(since))
        queried = True
        for row in results:
            # One unwritable person must not abandon the other thousands.
            # Count it, keep going, and let the summary show the damage.
            try:
                if row.Action == "UPDATE":
                    model.AddExtraValueDate(row.PeopleId, EV_NAME, row.CalcServiceDate)
                    updated += 1
                elif row.Action == "CLEAR":
                    model.DeleteExtraValue(row.PeopleId, EV_NAME)
                    cleared += 1
            except Exception as row_err:
                errors += 1
                if errors <= 3:
                    print "ServiceDate: could not update PeopleId {0}: {1}".format(
                        getattr(row, 'PeopleId', '?'), str(row_err))
    except Exception as e:
        print "ServiceDate {0} run FAILED: {1}".format(mode, str(e))

    # Only stamp the watermark when the query actually ran. If it did not, the
    # next run must re-cover this ground rather than skipping it. Row-level
    # errors are different: those people simply stay wrong until the next full
    # rescan picks them up, and advancing is still correct for everyone else.
    if queried:
        try:
            record_run(mode, updated, cleared,
                       since.strftime('%Y-%m-%d %H:%M:%S') if since else '')
        except Exception as state_err:
            print "ServiceDate: run completed but state not saved: {0}".format(str(state_err))

        print "ServiceDate {0} run ({1}): {2} updated, {3} cleared{4}".format(
            mode, why, updated, cleared,
            ", {0} errors".format(errors) if errors else "")


# ===========================================================================
# MANUAL MODE - Dashboard + Preload/Cleanup UI
# ===========================================================================
else:
    model.Header = "ServiceDate"

    # -----------------------------------------------------------------
    # Setup check -- reads THIS install's lookup tables so a new church
    # configures from real ids instead of guessing. The serve-type check
    # matters most: a wrong id there fails silently, crediting nobody.
    # -----------------------------------------------------------------
    setup_rows = ""
    setup_problem = False
    try:
        org_types = q.QuerySql("""
            SELECT ot.Id, ot.Description,
                   (SELECT COUNT(*) FROM Organizations o
                     WHERE o.OrganizationTypeId = ot.Id) AS OrgCount
            FROM lookup.OrganizationType ot ORDER BY ot.Id
        """)
        for t in org_types:
            chosen = int(t.Id) in [int(x) for x in SERVE_ORG_TYPE_IDS]
            mark = '&#10004; counted as serving' if chosen else ''
            style = 'background:#e8f5e9;font-weight:600;' if chosen else ''
            setup_rows += ('<tr style="{0}"><td>{1}</td><td>{2}</td>'
                           '<td align="right">{3}</td><td>{4}</td></tr>').format(
                style, t.Id, t.Description, t.OrgCount, mark)
    except Exception as e:
        setup_rows = '<tr><td colspan="4">Could not read lookup.OrganizationType: {0}</td></tr>'.format(str(e))

    # Does the configured serve type actually match any organizations?
    serve_org_count = 0
    try:
        if SERVE_ORG_TYPE_IDS:
            r = q.QuerySqlTop1("SELECT COUNT(*) AS n FROM Organizations WHERE OrganizationTypeId IN ({0})".format(
                ','.join(str(int(x)) for x in SERVE_ORG_TYPE_IDS)))
            serve_org_count = int(r.n) if r else 0
    except:
        serve_org_count = 0

    warn_html = ""
    if not SERVE_ORG_TYPE_IDS:
        setup_problem = True
        warn_html = ('<div style="background:#fdecea;border:1px solid #f5c6cb;padding:10px;'
                     'border-radius:4px;margin-bottom:10px;">'
                     '<b>SERVE_ORG_TYPE_IDS is empty.</b> Serving will not count toward '
                     'ServiceDate at all. Pick the row below that matches your serving '
                     'involvements and put its Id in the config block.</div>')
    elif serve_org_count == 0:
        setup_problem = True
        warn_html = ('<div style="background:#fdecea;border:1px solid #f5c6cb;padding:10px;'
                     'border-radius:4px;margin-bottom:10px;">'
                     '<b>SERVE_ORG_TYPE_IDS = {0} matches zero organizations on this install.</b> '
                     'Nobody will ever be credited for serving. Pick the correct Id from the '
                     'table below.</div>').format(SERVE_ORG_TYPE_IDS)

    # -----------------------------------------------------------------
    # Run status -- what the next morning batch will actually do
    # -----------------------------------------------------------------
    st = load_state() or {}
    next_mode, next_since, next_why = decide_mode()
    last_run_txt = st.get('last_run') or 'never'
    last_full_txt = st.get('last_full_run') or 'never'
    last_summary = ''
    if st.get('last_run'):
        last_summary = ' &mdash; {0} run, {1} updated, {2} cleared'.format(
            st.get('last_mode', '?'), st.get('last_updated', 0), st.get('last_cleared', 0))

    sched = sched_status()
    if not sched.get('readable'):
        sched_line = ('<span style="color:#a35c00;">Could not read '
                      + _SCHED_CONTENT_SLOT + ' to check.</span>')
    elif sched.get('installed'):
        # Show the name parsed out of the installed block, not our guess. If the
        # two disagree the block is calling something else and needs reinstalling.
        reg = sched.get('registered_name') or '(could not read the name)'
        sched_line = ('<span style="color:#1d6b2a;font-weight:600;">&#10004; Installed '
                      'in ' + _SCHED_CONTENT_SLOT + '</span> as <code>'
                      + reg
                      + '</code> '
                      '<button class="btn" style="background:#6c757d;color:#fff;'
                      'padding:2px 8px;font-size:0.85em;" onclick="schedUninstall()">Remove</button>')
    else:
        sched_line = ('<span style="color:#c62828;font-weight:600;">&#10007; Not in '
                      + _SCHED_CONTENT_SLOT + '</span> &mdash; nothing is keeping '
                      + EV_NAME + ' current. '
                      '<button class="btn btn-primary" style="padding:2px 10px;font-size:0.85em;" '
                      'onclick="schedInstall()">Add to Morning Batch</button>')
    if sched.get('outside'):
        sched_line += ('<br><span style="color:#a35c00;">Heads up: this script is also '
                       'called from ' + _SCHED_CONTENT_SLOT + ' outside the managed block. '
                       'Installing here as well would run it twice.</span>')
    if sched.get('stray'):
        # An earlier build of this page installed into ScheduledTasks, which is
        # not the morning batch. That block looks installed and never runs, so
        # it carries its own Remove button - when we are NOT installed in the
        # real slot there is no other Remove button on the page to point at.
        sched_line += ('<br><span style="color:#c62828;">Found our block in <b>'
                       + ', '.join(sched.get('stray')) + '</b>, which does not run '
                       'with the morning batch. </span>'
                       '<button class="btn" style="background:#c62828;color:#fff;'
                       'padding:2px 8px;font-size:0.85em;" onclick="schedUninstall()">'
                       'Clean Up</button>')

    status_html = (
        '<div style="background:#eef5fb;border:1px solid #cfe3ff;padding:10px;'
        'border-radius:4px;margin-bottom:10px;font-size:0.92em;">'
        '<b>Morning batch status</b><br>'
        'Last run: <b>{0}</b>{1}<br>'
        'Last full rescan: <b>{2}</b><br>'
        'Next run will be: <b>{3}</b> ({4})<br>'
        'Morning batch: {8}<br>'
        '<span style="color:#555;">Trigger: <code>Data.{5} = "true"</code> then '
        '<code>model.CallScript("<span class="sdName">{9}</span>")</code>. '
        # Correct the name inline, during parse, so the server's guess is never
        # painted. sdFixScriptName() in model.Script does this too, but it lands
        # after first paint and the admin copies this line into ScheduledTasks -
        # a flash of the wrong name is a flash of a wrong instruction. Braces are
        # doubled because this string goes through .format().
        '<script>(function()'
        '{{var p=location.pathname||"",'
        'm=p.match(/\/PyScript(?:Form)?\/([^\/?#]+)/i),'
        'n=m&&m[1]?decodeURIComponent(m[1]):"";'
        'if(!n)return;'
        'var s=document.getElementsByClassName("sdName");'
        'for(var i=0;i<s.length;i++)s[i].textContent=n;}})();</script>'
        'Incremental runs re-check the last {6} day(s) as overlap; a full rescan '
        'happens at least every {7} day(s).</span>'
        '<div id="schedMsg" style="margin-top:6px;"></div>'
        '</div>').format(
            last_run_txt, last_summary, last_full_txt, next_mode.upper(), next_why,
            BATCH_TRIGGER, INCREMENTAL_OVERLAP_DAYS, FULL_RESCAN_EVERY_DAYS,
            sched_line, sched.get('script_name', _DEFAULT_SCRIPT_NAME))

    # -----------------------------------------------------------------
    # Search Builder scripts. The Extra Value is only useful if staff can
    # query it, so surface the In SQL List scripts next to the batch
    # status rather than leaving people to create them by hand.
    # -----------------------------------------------------------------
    sq = sqllist_status()
    chips = []
    for it in sq['items']:
        if it['state'] == 'ok':
            colour, mark, note = '#1d6b2a', '&#10004;', ''
        elif it['state'] == 'differs':
            colour, mark, note = '#a35c00', '&#9679;', ' (edited)'
        else:
            colour, mark, note = '#c62828', '&#10007;', ' (missing)'
        chips.append('<span style="color:' + colour + ';font-weight:600;'
                     'margin-right:12px;white-space:nowrap;">' + mark +
                     ' <code>' + it['name'] + '</code>' + note + '</span>')

    if not sq['readable']:
        sql_action = ('<span style="color:#a35c00;">Could not read Special Content '
                      'to check these.</span>')
    elif sq['missing'] or sq['differs']:
        label = 'Create Search Builder Scripts' if not sq['differs'] else \
                'Create Missing Scripts'
        sql_action = ('<button class="btn btn-primary" style="padding:2px 10px;'
                      'font-size:0.85em;" onclick="sqlListInstall(false)">' + label + '</button>')
        if sq['differs']:
            sql_action += (' <button class="btn" style="background:#6c757d;color:#fff;'
                           'padding:2px 8px;font-size:0.85em;margin-left:6px;" '
                           'onclick="sqlListInstall(true)">Overwrite Edited</button>')
    else:
        sql_action = ('<span style="color:#1d6b2a;">All present. Search Builder &rarr; '
                      '<b>In SQL List</b> &rarr; pick one.</span>')

    status_html += (
        '<div style="background:#f4f7f2;border:1px solid #d6e4cd;padding:10px;'
        'border-radius:4px;margin-bottom:10px;font-size:0.92em;">'
        '<b>Search Builder scripts</b><br>'
        '<span style="color:#555;">Saved SQL scripts staff pick from '
        '<b>In SQL List</b> to pull everyone who served within a window.</span><br>'
        '<div style="margin:6px 0;">' + ''.join(chips) + '</div>'
        + sql_action +
        '<div id="sqlListMsg" style="margin-top:6px;"></div>'
        '</div>')

    # -----------------------------------------------------------------
    # Settings modal. Checkbox list for the org types so nobody has to
    # know an id; plain fields for the rest. Built from CONFIG_DEFAULTS so
    # the form and the loader cannot drift apart.
    # -----------------------------------------------------------------
    type_checks = ""
    try:
        for t in q.QuerySql("""
            SELECT ot.Id, ot.Description,
                   (SELECT COUNT(*) FROM Organizations o
                     WHERE o.OrganizationTypeId = ot.Id) AS OrgCount
            FROM lookup.OrganizationType ot ORDER BY ot.Description
        """):
            checked = 'checked' if int(t.Id) in [int(x) for x in SERVE_ORG_TYPE_IDS] else ''
            type_checks += (
                '<label style="display:inline-block;width:32%;min-width:190px;margin:2px 0;">'
                '<input type="checkbox" class="serveType" value="{0}" {1}> '
                '{2} <span style="color:#888;">({3} orgs)</span></label>').format(
                    t.Id, checked, t.Description, t.OrgCount)
    except Exception as e:
        type_checks = '<i>Could not read lookup.OrganizationType: {0}</i>'.format(str(e))

    contrib_rows = ""
    try:
        for ct in q.QuerySql("SELECT Id, Description FROM lookup.ContributionType ORDER BY Id"):
            contrib_rows += '<option value="{0}">{0} &mdash; {1}</option>'.format(
                ct.Id, ct.Description)
    except:
        contrib_rows = ""

    settings_html = (
        '<div id="cfgModal" style="display:none;position:fixed;top:0;left:0;right:0;bottom:0;'
        'background:rgba(0,0,0,0.5);z-index:9999;overflow:auto;">'
        '<div style="background:#fff;max-width:760px;margin:40px auto;padding:20px;'
        'border-radius:6px;">'
        '<h3 style="margin-top:0;">ServiceDate settings</h3>'
        '<p style="color:#666;font-size:0.9em;">Saved to Special Content, so this survives '
        'script updates. After changing anything here, run <b>Force Full Rescan</b> &mdash; '
        'an incremental run only looks at new activity and will not revisit everyone.</p>'

        '<fieldset style="margin-bottom:12px;"><legend><b>What counts as serving</b></legend>'
        + type_checks +
        '<p style="color:#a35c00;font-size:0.85em;margin:6px 0 0;">Tick nothing and serving '
        'stops counting entirely.</p></fieldset>'

        '<fieldset style="margin-bottom:12px;"><legend><b>What counts as giving</b></legend>'
        '<label>Exclude contribution type ids '
        '<input type="text" id="cfg_EXCLUDE_CONTRIBUTION_TYPE_IDS" value="{0}" size="24"></label>'
        '<div style="color:#888;font-size:0.85em;margin:4px 0 8px;">Comma separated. '
        'Your types: <select onchange="if(this.value)addExcl(this.value)">'
        '<option value="">(look up an id)</option>' + contrib_rows + '</select></div>'
        '<label>Contribution status id (posted) '
        '<input type="text" id="cfg_CONTRIBUTION_STATUS_ID" value="{1}" size="4"></label><br>'
        '<label><input type="checkbox" id="cfg_CREDIT_SPOUSE_FOR_GIVING" {2}> '
        'Give a spouse credit for household giving</label>'
        '</fieldset>'

        '<fieldset style="margin-bottom:12px;"><legend><b>Sources &amp; scope</b></legend>'
        '<label><input type="checkbox" id="cfg_INCLUDE_MISSION_TRIPS" {3}> '
        'Count mission trips</label><br>'
        '<label>Full rescan looks back '
        '<input type="text" id="cfg_SCOPE_DAYS" value="{4}" size="5"> days</label><br>'
        '<label>Extra Value name '
        '<input type="text" id="cfg_EV_NAME" value="{5}" size="18"></label>'
        '</fieldset>'

        '<fieldset style="margin-bottom:12px;"><legend><b>Morning batch</b></legend>'
        '<label>Incremental overlap '
        '<input type="text" id="cfg_INCREMENTAL_OVERLAP_DAYS" value="{6}" size="4"> days</label>'
        '<span style="color:#888;font-size:0.85em;"> re-checked each run to catch back-dated '
        'entries</span><br>'
        '<label>Force a full rescan every '
        '<input type="text" id="cfg_FULL_RESCAN_EVERY_DAYS" value="{7}" size="4"> days</label>'
        '<span style="color:#888;font-size:0.85em;"> only a full pass can clear people who '
        'aged out</span><br>'
        '<label>Batch trigger flag '
        '<input type="text" id="cfg_BATCH_TRIGGER" value="{8}" size="20"></label>'
        '<span style="color:#888;font-size:0.85em;"> must match the '
        '<code>Data.&lt;name&gt; = "true"</code> line</span>'
        '</fieldset>'

        '<div id="cfgMsg" style="margin:8px 0;"></div>'
        '<button class="btn btn-primary" onclick="saveConfig()">Save settings</button> '
        '<button class="btn" style="background:#6c757d;color:#fff;" '
        'onclick="document.getElementById(\'cfgModal\').style.display=\'none\'">Cancel</button>'
        '</div></div>').format(
            ','.join(str(x) for x in EXCLUDE_CONTRIBUTION_TYPE_IDS),
            CONTRIBUTION_STATUS_ID,
            'checked' if CREDIT_SPOUSE_FOR_GIVING else '',
            'checked' if INCLUDE_MISSION_TRIPS else '',
            SCOPE_DAYS, EV_NAME,
            INCREMENTAL_OVERLAP_DAYS, FULL_RESCAN_EVERY_DAYS, BATCH_TRIGGER)

    setup_html = (
        '<details style="margin-bottom:12px;" {0}>'
        '<summary style="cursor:pointer;font-weight:600;">Setup check &mdash; '
        'configure this for your church</summary>'
        '<div style="padding:10px 0;">{1}'
        '<p style="margin:6px 0;">Current config: '
        'serve types <code>{2}</code> ({3} orgs) &middot; '
        'mission trips <code>{4}</code> &middot; '
        'excluded contribution types <code>{5}</code> &middot; '
        'spouse credit <code>{6}</code> &middot; '
        'lookback <code>{7}</code> days</p>'
        '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;font-size:0.9em;">'
        '<tr style="background:#f5f5f5;"><th>Id</th><th align="left">Organization Type</th>'
        '<th>Orgs</th><th></th></tr>{8}</table>'
        '</div></details>').format(
            'open' if setup_problem else '',
            warn_html,
            SERVE_ORG_TYPE_IDS, serve_org_count,
            INCLUDE_MISSION_TRIPS, EXCLUDE_CONTRIBUTION_TYPE_IDS,
            CREDIT_SPOUSE_FOR_GIVING, SCOPE_DAYS,
            setup_rows)


    results = q.QuerySql(get_sql())

    # Check for legacy EVs to clean up
    legacy_results = q.QuerySql(get_cleanup_sql())
    legacy_count = 0
    for row in legacy_results:
        legacy_count += 1

    total_people = 0
    gave_count = 0
    mission_count = 0
    serve_count = 0
    update_count = 0
    clear_count = 0
    ok_count = 0
    rows_html = ""

    for row in results:
        total_people += 1
        if row.LastGaveDate:
            gave_count += 1
        if row.LastMissionDate:
            mission_count += 1
        if row.LastServeDate:
            serve_count += 1

        if row.Action == "UPDATE":
            update_count += 1
        elif row.Action == "CLEAR":
            clear_count += 1
        else:
            ok_count += 1

        # Format dates
        calc_display = row.CalcServiceDate.ToString("MM/dd/yyyy") if row.CalcServiceDate else "N/A"
        existing_display = row.ExistingServiceDate.ToString("MM/dd/yyyy") if row.ExistingServiceDate else "-"
        days_since = row.DaysSince if row.DaysSince is not None else ""

        gave_display = row.LastGaveDate.ToString("MM/dd/yyyy") if row.LastGaveDate else "-"
        mission_display = row.LastMissionDate.ToString("MM/dd/yyyy") if row.LastMissionDate else "-"
        serve_display = row.LastServeDate.ToString("MM/dd/yyyy") if row.LastServeDate else "-"

        driver = row.Driver if row.Driver else ""
        driver_colors = {"Gave": "#27ae60", "Mission": "#2980b9", "Serve": "#8e44ad"}
        driver_color = driver_colors.get(driver, "#333")

        action_colors = {"UPDATE": "#e67e22", "CLEAR": "#e74c3c", "OK": "#27ae60"}
        action_color = action_colors.get(row.Action, "#333")

        rows_html += '''<tr>
            <td><a href="/Person2/{0}" target="_blank">{1}</a></td>
            <td>{2}</td>
            <td>{3}</td>
            <td>{4}</td>
            <td>{5}</td>
            <td>{6}</td>
            <td style="font-weight:bold">{7}</td>
            <td>{8}</td>
            <td style="font-weight:bold;color:{12}">{9}</td>
            <td>{10}</td>
            <td><span style="background:{13};color:#fff;padding:2px 8px;border-radius:3px;font-size:11px">{11}</span></td>
        </tr>'''.format(
            row.PeopleId,          # 0
            row.PersonName or "",  # 1
            row.MemberStatus or "",  # 2
            row.FamilyPosition or "",  # 3
            gave_display,          # 4
            mission_display,       # 5
            serve_display,         # 6
            calc_display,          # 7
            days_since,            # 8
            driver,                # 9
            existing_display,      # 10
            row.Action,            # 11
            driver_color,          # 12
            action_color           # 13
        )

    # Legacy cleanup banner
    legacy_banner = ""
    if legacy_count > 0:
        legacy_banner = '''
        <div style="background:#ffeaa7;border:1px solid #fdcb6e;padding:12px 20px;border-radius:6px;margin:15px 0;display:flex;align-items:center;gap:15px;flex-wrap:wrap">
            <div>
                <strong>Legacy Cleanup:</strong> Found <strong>{0}</strong> people with old "ServiceDays" (integer) Extra Value.
                This field is no longer used and should be removed.
            </div>
            <button class="btn" style="background:#e74c3c;color:#fff" id="btnCleanup" onclick="startCleanup()">
                Remove Legacy EV ({0} records)
            </button>
            <div id="cleanupStatus" style="font-size:13px;color:#555"></div>
        </div>'''.format(legacy_count)

    model.Form = '''<html>
<head>
<style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 20px; }}
    h2 {{ color: #2c3e50; margin-bottom: 5px; }}
    .subtitle {{ color: #666; margin-bottom: 20px; }}
    .summary {{ display: flex; gap: 15px; margin: 20px 0; flex-wrap: wrap; }}
    .stat-card {{ background: #f8f9fa; border-left: 4px solid #333; padding: 12px 18px; border-radius: 4px; min-width: 130px; }}
    .stat-card.gave {{ border-color: #27ae60; }}
    .stat-card.mission {{ border-color: #2980b9; }}
    .stat-card.serve {{ border-color: #8e44ad; }}
    .stat-card.total {{ border-color: #e67e22; }}
    .stat-card.update {{ border-color: #e67e22; }}
    .stat-card.clear {{ border-color: #e74c3c; }}
    .stat-card.ok {{ border-color: #27ae60; }}
    .stat-card h4 {{ margin: 0 0 4px 0; color: #666; font-size: 11px; text-transform: uppercase; }}
    .stat-card .number {{ font-size: 24px; font-weight: bold; color: #2c3e50; }}
    table {{ border-collapse: collapse; width: 100%; margin: 20px 0; font-size: 13px; }}
    th {{ background: #2c3e50; color: white; padding: 10px 12px; text-align: left; position: sticky; top: 0; z-index: 1; }}
    td {{ border-bottom: 1px solid #e0e0e0; padding: 8px 12px; }}
    tr:hover {{ background: #f5f5f5; }}
    a {{ color: #2980b9; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .controls {{ background: #f0f4f8; padding: 15px 20px; border-radius: 6px; margin: 15px 0; display: flex; align-items: center; gap: 15px; flex-wrap: wrap; }}
    .btn {{ padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 600; }}
    .btn-primary {{ background: #2980b9; color: white; }}
    .btn-primary:hover {{ background: #2471a3; }}
    .btn-primary:disabled {{ background: #bdc3c7; cursor: not-allowed; }}
    .progress-bar {{ flex: 1; min-width: 200px; }}
    .progress-outer {{ background: #ddd; border-radius: 10px; height: 20px; overflow: hidden; }}
    .progress-inner {{ background: #27ae60; height: 100%; transition: width 0.3s; border-radius: 10px; }}
    #status {{ font-size: 13px; color: #555; }}
</style>
</head>
<body>
<h2>ServiceDate Dashboard</h2>
<p class="subtitle">Scope: {6} days | EV: {11} | Generated: {0}</p>

{12}
{13}
{14}
{10}

<div class="summary">
    <div class="stat-card total">
        <h4>Total People</h4>
        <div class="number">{1}</div>
    </div>
    <div class="stat-card gave">
        <h4>Gave</h4>
        <div class="number">{2}</div>
    </div>
    <div class="stat-card mission">
        <h4>Mission Trip</h4>
        <div class="number">{3}</div>
    </div>
    <div class="stat-card serve">
        <h4>Served</h4>
        <div class="number">{4}</div>
    </div>
</div>

<div class="summary">
    <div class="stat-card update">
        <h4>Needs Update</h4>
        <div class="number">{7}</div>
    </div>
    <div class="stat-card clear">
        <h4>Needs Clear</h4>
        <div class="number">{8}</div>
    </div>
    <div class="stat-card ok">
        <h4>Up to Date</h4>
        <div class="number">{9}</div>
    </div>
</div>

<div class="controls">
    <button class="btn btn-primary" id="btnPreload" onclick="startPreload()">
        Run Preload ({7} updates, {8} clears)
    </button>
    <button class="btn" style="background:#6c757d;color:#fff;margin-left:8px;"
            id="btnFullRescan" onclick="startFullRescan()">
        Force Full Rescan
    </button>
    <button class="btn" style="background:#34495e;color:#fff;margin-left:8px;"
            onclick="document.getElementById('cfgModal').style.display='block'">
        Settings
    </button>
    <div style="font-size:0.85em;color:#666;margin-top:6px;">
        <b>Preload</b> populates {11} for the first time.
        <b>Force Full Rescan</b> does the same work again from scratch &mdash; run it after
        you change any setting in the config block (especially the serve types), because an
        incremental run only looks at new activity and would never revisit everyone.
        Both run in batches, so neither will time out, and both reset the
        weekly-full-rescan clock.
    </div>
    <div class="progress-bar">
        <div class="progress-outer">
            <div class="progress-inner" id="progressBar" style="width:0%"></div>
        </div>
    </div>
    <div id="status">Ready</div>
</div>

<table>
    <thead>
        <tr>
            <th>Person</th>
            <th>Status</th>
            <th>Family Position</th>
            <th style="color:#27ae60">Last Gave</th>
            <th style="color:#2980b9">Last Mission</th>
            <th style="color:#8e44ad">Last Served</th>
            <th>Calculated Date</th>
            <th>Days Since</th>
            <th>Driven By</th>
            <th>Existing EV</th>
            <th>Action</th>
        </tr>
    </thead>
    <tbody>
        {5}
    </tbody>
</table>
</body>
</html>'''.format(
        model.DateTime.Now.ToString("MM/dd/yyyy 'at' hh:mm tt"),  # 0
        total_people,       # 1
        gave_count,         # 2
        mission_count,      # 3
        serve_count,        # 4
        rows_html,          # 5
        SCOPE_DAYS,         # 6
        update_count,       # 7
        clear_count,        # 8
        ok_count,           # 9
        legacy_banner,      # 10
        EV_NAME,            # 11
        status_html,        # 12  morning-batch status
        setup_html,         # 13  setup check for a new install
        settings_html       # 14  settings modal (hidden until opened)
    )

    model.Script = '''
var totalProcessed = 0;
var initialTotal = 0;

// Force Full Rescan reuses the preload loop on purpose: a preload IS a full
// pass, and that path is already batched so it cannot time out on a large
// church. The only difference is the confirm and the wording.
// Everything below posts to this same script. script_name rides along so the
// server can register the RIGHT name in ScheduledTasks even when the admin has
// renamed the script.
function sdScriptName() {
    var path = window.location.pathname || '';
    // Preferred: pull the name straight out of the TouchPoint route, so a
    // trailing slash or an extra segment cannot throw us off.
    var m = path.match(/\/PyScript(?:Form)?\/([^\/?#]+)/i);
    if (m && m[1]) { return decodeURIComponent(m[1]); }
    // Fallback: last non-empty segment.
    var parts = path.split('/');
    while (parts.length && !parts[parts.length - 1]) { parts.pop(); }
    return parts.length ? parts[parts.length - 1].split('?')[0] : '';
}

// TouchPoint does not hand the script's own name to Python, so the server
// rendered its best guess (the default) into every .sdName span. The browser
// DOES know, from the URL, so correct those and re-run the outside-reference
// check with the real name. Without this the page tells an admin to register
// "TPxi_ServiceDays" even when they installed it as something else.
// Note: the "Installed as X" banner deliberately has NO .sdName span - that
// name is read back out of ScheduledTasks and is the truth, so if it differs
// from the current URL the admin needs to see the difference, not have it
// papered over.
function sdFixScriptName() {
    var real = sdScriptName();
    if (!real) return;
    var spans = document.getElementsByClassName("sdName");
    for (var i = 0; i < spans.length; i++) { spans[i].textContent = real; }

    // The installed/not-installed banner is name-independent (it looks for our
    // marker), but the "also called outside the block" warning is not, so ask
    // the server again now that we can tell it the truth.
    $.ajax({
        url: window.location.pathname, type: "POST",
        data: { action: "sched_status", script_name: real },
        success: function(resp) {
            try {
                var d = JSON.parse(resp);
                if (d && d.outside) {
                    var msg = document.getElementById("schedMsg");
                    if (msg && msg.innerHTML.indexOf("outside") < 0) {
                        msg.innerHTML = '<span style="color:#a35c00;">Heads up: ' + real +
                            ' is also called from ScheduledTasks outside the managed ' +
                            'block. Installing here as well would run it twice.</span>';
                    }
                }
            } catch(e) {}
        }
    });
}

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", sdFixScriptName);
} else {
    sdFixScriptName();
}

function addExcl(v) {
    var f = document.getElementById("cfg_EXCLUDE_CONTRIBUTION_TYPE_IDS");
    var cur = f.value.split(',').map(function(x){return x.trim();}).filter(function(x){return x;});
    if (cur.indexOf(v) < 0) { cur.push(v); f.value = cur.join(','); }
}

function saveConfig() {
    var types = [];
    var boxes = document.getElementsByClassName("serveType");
    for (var i = 0; i < boxes.length; i++) {
        if (boxes[i].checked) types.push(boxes[i].value);
    }
    var payload = {
        action: "save_config",
        script_name: sdScriptName(),
        SERVE_ORG_TYPE_IDS: types.join(','),
        EXCLUDE_CONTRIBUTION_TYPE_IDS: document.getElementById("cfg_EXCLUDE_CONTRIBUTION_TYPE_IDS").value,
        CONTRIBUTION_STATUS_ID: document.getElementById("cfg_CONTRIBUTION_STATUS_ID").value,
        SCOPE_DAYS: document.getElementById("cfg_SCOPE_DAYS").value,
        EV_NAME: document.getElementById("cfg_EV_NAME").value,
        INCREMENTAL_OVERLAP_DAYS: document.getElementById("cfg_INCREMENTAL_OVERLAP_DAYS").value,
        FULL_RESCAN_EVERY_DAYS: document.getElementById("cfg_FULL_RESCAN_EVERY_DAYS").value,
        BATCH_TRIGGER: document.getElementById("cfg_BATCH_TRIGGER").value
    };
    // Unticked checkboxes must be sent explicitly as false, otherwise the
    // server cannot tell "off" from "not submitted".
    if (document.getElementById("cfg_CREDIT_SPOUSE_FOR_GIVING").checked) payload.CREDIT_SPOUSE_FOR_GIVING = "true";
    if (document.getElementById("cfg_INCLUDE_MISSION_TRIPS").checked) payload.INCLUDE_MISSION_TRIPS = "true";

    document.getElementById("cfgMsg").innerHTML = "Saving...";
    $.ajax({
        url: window.location.pathname, type: "POST", data: payload,
        success: function(resp) {
            try {
                var d = JSON.parse(resp);
                document.getElementById("cfgMsg").innerHTML = d.success
                    ? '<span style="color:green">' + d.message + ' Reloading...</span>'
                    : '<span style="color:red">' + d.message + '</span>';
                if (d.success) setTimeout(function(){ window.location.reload(); }, 1200);
            } catch(e) {
                document.getElementById("cfgMsg").innerHTML =
                    '<span style="color:red">Unexpected response: ' + e.message + '</span>';
            }
        },
        error: function(xhr) {
            document.getElementById("cfgMsg").innerHTML =
                '<span style="color:red">Save failed: ' + xhr.status + '</span>';
        }
    });
}

// Creating the In SQL List scripts. Overwrite is a separate, confirmed button
// because replacing a script someone has tuned by hand is not recoverable from
// this page.
function sqlListInstall(overwrite) {
    if (overwrite && !confirm("Replace the edited script(s) with the standard " +
                              "version?\\n\\nAny changes made to them in Special " +
                              "Content will be lost.")) {
        return;
    }
    document.getElementById("sqlListMsg").innerHTML = "Working...";
    $.ajax({
        url: window.location.pathname, type: "POST",
        data: { action: "sqllist_install", overwrite: overwrite ? "true" : "false",
                script_name: sdScriptName() },
        success: function(resp) {
            try {
                var d = JSON.parse(resp);
                document.getElementById("sqlListMsg").innerHTML = d.success
                    ? '<span style="color:green">' + d.message + ' Reloading...</span>'
                    : '<span style="color:red">' + d.message + '</span>';
                if (d.success) setTimeout(function(){ window.location.reload(); }, 1400);
            } catch(e) {
                document.getElementById("sqlListMsg").innerHTML =
                    '<span style="color:red">Unexpected response: ' + e.message + '</span>';
            }
        },
        error: function(xhr) {
            document.getElementById("sqlListMsg").innerHTML =
                '<span style="color:red">Failed: ' + xhr.status + '</span>';
        }
    });
}

function schedCall(act, confirmMsg) {
    if (confirmMsg && !confirm(confirmMsg)) return;
    document.getElementById("schedMsg").innerHTML = "Working...";
    $.ajax({
        url: window.location.pathname, type: "POST",
        data: { action: act, script_name: sdScriptName() },
        success: function(resp) {
            try {
                var d = JSON.parse(resp);
                document.getElementById("schedMsg").innerHTML = d.success
                    ? '<span style="color:green">' + d.message + ' Reloading...</span>'
                    : '<span style="color:red">' + d.message + '</span>';
                if (d.success) setTimeout(function(){ window.location.reload(); }, 1200);
            } catch(e) {
                document.getElementById("schedMsg").innerHTML =
                    '<span style="color:red">Unexpected response: ' + e.message + '</span>';
            }
        },
        error: function(xhr) {
            document.getElementById("schedMsg").innerHTML =
                '<span style="color:red">Failed: ' + xhr.status + '</span>';
        }
    });
}

function schedInstall()   { schedCall("sched_install"); }
function schedUninstall() { schedCall("sched_uninstall",
    "Remove ServiceDate from the morning batch? __EV_NAME__ will stop updating."); }

function startFullRescan() {
    if (!confirm("Recompute __EV_NAME__ for everyone from scratch?\\n\\n" +
                 "Use this after changing the configuration. It runs in batches " +
                 "and may take a few minutes on a large database.")) {
        return;
    }
    var full = document.getElementById("btnFullRescan");
    full.disabled = true;
    full.textContent = "Rescanning...";
    startPreload();
}

function startPreload() {
    var btn = document.getElementById("btnPreload");
    btn.disabled = true;
    btn.textContent = "Running...";
    totalProcessed = 0;
    initialTotal = 0;
    document.getElementById("status").textContent = "Starting...";
    runBatch();
}

function runBatch() {
    $.ajax({
        url: window.location.pathname,
        type: "POST",
        data: { action: "preload" },
        success: function(resp) {
            try {
                var data = JSON.parse(resp);
                if (initialTotal === 0) initialTotal = data.total_found;
                totalProcessed += data.processed;
                var pct = Math.round((totalProcessed / initialTotal) * 100);
                document.getElementById("progressBar").style.width = pct + "%";
                document.getElementById("status").textContent =
                    data.updated + " updated, " + data.cleared + " cleared this batch | " +
                    totalProcessed + " done, " + data.remaining + " remaining";

                if (data.has_more) {
                    runBatch();
                } else {
                    document.getElementById("progressBar").style.width = "100%";
                    document.getElementById("status").textContent =
                        "Complete! " + totalProcessed + " total processed.";
                    document.getElementById("btnPreload").textContent = "Done - Refresh to verify";
                    var fullBtn = document.getElementById("btnFullRescan");
                    if (fullBtn) { fullBtn.textContent = "Done"; }
                }
            } catch(e) {
                document.getElementById("status").textContent = "Error: " + e.message;
                document.getElementById("btnPreload").disabled = false;
                document.getElementById("btnPreload").textContent = "Retry";
                var fb = document.getElementById("btnFullRescan");
                if (fb) { fb.disabled = false; fb.textContent = "Force Full Rescan"; }
            }
        },
        error: function(xhr) {
            document.getElementById("status").textContent = "Request failed: " + xhr.status + ". " + totalProcessed + " already done.";
            document.getElementById("btnPreload").disabled = false;
            document.getElementById("btnPreload").textContent = "Retry (" + totalProcessed + " done so far)";
            var fb = document.getElementById("btnFullRescan");
            if (fb) { fb.disabled = false; fb.textContent = "Force Full Rescan"; }
        }
    });
}

function startCleanup() {
    var btn = document.getElementById("btnCleanup");
    btn.disabled = true;
    btn.textContent = "Removing...";
    runCleanupBatch();
}

function runCleanupBatch() {
    $.ajax({
        url: window.location.pathname,
        type: "POST",
        data: { action: "cleanup" },
        success: function(resp) {
            try {
                var data = JSON.parse(resp);
                document.getElementById("cleanupStatus").textContent =
                    "Removed " + data.removed + " this batch, " + data.remaining + " remaining";

                if (data.has_more) {
                    runCleanupBatch();
                } else {
                    document.getElementById("cleanupStatus").textContent =
                        "Done! All legacy records removed.";
                    document.getElementById("btnCleanup").textContent = "Done";
                }
            } catch(e) {
                document.getElementById("cleanupStatus").textContent = "Error: " + e.message;
                document.getElementById("btnCleanup").disabled = false;
                document.getElementById("btnCleanup").textContent = "Retry";
            }
        },
        error: function(xhr) {
            document.getElementById("cleanupStatus").textContent = "Failed: " + xhr.status;
            document.getElementById("btnCleanup").disabled = false;
            document.getElementById("btnCleanup").textContent = "Retry";
        }
    });
}
'''.replace('__EV_NAME__', EV_NAME)
