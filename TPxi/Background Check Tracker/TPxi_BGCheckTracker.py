#Roles=Edit

"""
TPxi Background Check Tracker
==============================
A focused utility for triaging in-flight background checks. Extracted from
the Compliance Dashboard's Pending BG tab and enhanced with on-demand
MinistrySafe API lookups so you can see the LIVE status alongside the
TouchPoint cached status — solves the common "MS says Pending but TP
says Overdue" sync-drift problem.

What it does:
  * Lists every Pending background check (TouchPoint cache view)
  * Shows the live MinistrySafe status next to it on demand (per row or all)
  * Lets you resend the applicant invite for an "Awaiting Applicant" check
  * Direct links to the TouchPoint person profile and the MS report URL
  * Caches MS API results for 5 minutes to avoid hammering their API
  * Settings tab to enter / rotate the MinistrySafe API token with
    step-by-step instructions for finding it

What it does NOT do:
  * Write to TouchPoint's BackgroundChecks table. The TP cache is updated
    by TouchPoint's normal webhook / morning batch path. This utility
    just gives you visibility while you wait for that to catch up.

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

--Upload Instructions Start--
1. Admin -> Advanced -> Special Content -> Python -> New Python Script File
2. Name: TPxi_BGCheckTracker
3. Paste this entire file and Save
4. Add to CustomReports (optional) with role="Edit" if you want it in
   the report menu
5. Open Settings tab in the running script and paste your MinistrySafe
   API token (instructions on that tab walk through where to get it)
--Upload Instructions End--
"""

import json
import datetime
import time
import urllib2

model.Header = 'BG Check Tracker'

# --- Version / Auto-update -------------------------------------------------
# Bump APP_VERSION on every release published to the DisplayCache manifest.
# See TPxi/AutoUpdate/README.md for the mechanism.
APP_VERSION = '1.0.1'
DC_SCRIPT_ID = 'TPxi_BGCheckTracker'
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'


def get_script_name():
    """Detect the actual PythonContent slot this script was installed as."""
    try:
        if hasattr(Data, 'script_name') and Data.script_name:
            sn = str(Data.script_name).strip()
            if sn:
                return sn
    except:
        pass
    try:
        import re
        url = str(getattr(model, 'URL', '') or '')
        m = re.search(r'/PyScript(?:Form)?/([^/?#&]+)', url)
        if m:
            return m.group(1)
    except:
        pass
    return DC_SCRIPT_ID


# ============================================================================
# JSON safety (manual encoder — see TPxi/AutoUpdate/README.md and CLAUDE.md
# under "Common Gotchas #9". IronPython's json.dumps blows up on non-ASCII
# at the .NET boundary, even with ensure_ascii=True. Walking the structure
# ourselves and emitting \uXXXX for every codepoint >= 0x7F sidesteps it.)
# ============================================================================
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
        if   code == 0x22: parts.append('\\"')
        elif code == 0x5C: parts.append('\\\\')
        elif code == 0x08: parts.append('\\b')
        elif code == 0x0C: parts.append('\\f')
        elif code == 0x0A: parts.append('\\n')
        elif code == 0x0D: parts.append('\\r')
        elif code == 0x09: parts.append('\\t')
        elif code < 0x20 or code >= 0x7F:
            parts.append('\\u%04x' % code)
        else:
            parts.append(chr(code))
    parts.append('"')
    return ''.join(parts)


def _json_encode(obj):
    if obj is None: return 'null'
    if obj is True: return 'true'
    if obj is False: return 'false'
    if isinstance(obj, bool):
        return 'true' if obj else 'false'
    if isinstance(obj, (int, long)):
        return str(obj)
    if isinstance(obj, float):
        if obj != obj:
            return 'null'
        return repr(obj)
    if isinstance(obj, (str, unicode)):
        return _json_escape_string(obj)
    if isinstance(obj, datetime.datetime):
        return _json_escape_string(obj.isoformat())
    if isinstance(obj, datetime.date):
        return _json_escape_string(obj.isoformat())
    if isinstance(obj, dict):
        items = []
        for k, v in obj.items():
            key = k if isinstance(k, (str, unicode)) else unicode(k)
            items.append(_json_escape_string(key) + ':' + _json_encode(v))
        return '{' + ','.join(items) + '}'
    if isinstance(obj, (list, tuple)):
        return '[' + ','.join(_json_encode(x) for x in obj) + ']'
    return _json_escape_string(unicode(obj))


def safe_json(obj):
    try:
        return _json_encode(obj)
    except Exception as _exc:
        try:
            return '{"success":false,"message":' + _json_escape_string(
                'safe_json failed: ' + repr(_exc)) + '}'
        except Exception:
            return '{"success":false,"message":"safe_json catastrophic failure"}'


# ============================================================================
# Constants + Settings
# ============================================================================
MS_API_BASE = "https://safetysystem.ministrysafe.com/api/v2"
SETTING_API_TOKEN = "MinistrySafeApiToken"
CACHE_CONTENT_NAME = "TPxi_BGCheckTracker_Cache"
CACHE_TTL_SECONDS = 300  # 5-minute freshness on MS responses


def get_ms_api_token():
    """Read the MinistrySafe API token from TouchPoint Settings. Empty
    string when not configured — every API call short-circuits with a
    helpful error in that case."""
    try:
        val = model.Setting(SETTING_API_TOKEN, "")
        return str(val or "").strip()
    except Exception:
        return ""


def set_ms_api_token(token):
    """Persist the token to the Setting table."""
    model.SetSetting(SETTING_API_TOKEN, str(token or "").strip())


# ============================================================================
# Cache helpers — keyed by report_id, time-bound to CACHE_TTL_SECONDS
# ============================================================================
def _now_ts():
    return int(time.time())


def load_cache():
    """Returns a dict { report_id: {ts: int, status: str, raw: dict, error: str} }.
    Missing / corrupt blob returns empty dict."""
    try:
        raw = model.TextContent(CACHE_CONTENT_NAME)
        if raw:
            return json.loads(raw) or {}
    except Exception:
        pass
    return {}


def save_cache(cache):
    """Persist the cache blob. Failures are silent — we'd rather lose
    cache than crash the page."""
    try:
        model.WriteContentText(CACHE_CONTENT_NAME, safe_json(cache), "")
    except Exception:
        pass


def get_cached(report_id):
    """Return cached entry IF still fresh, else None."""
    if not report_id:
        return None
    cache = load_cache()
    entry = cache.get(str(report_id))
    if not entry:
        return None
    age = _now_ts() - int(entry.get('ts', 0))
    if age > CACHE_TTL_SECONDS:
        return None
    return entry


def put_cached(report_id, status, raw=None, error=None):
    if not report_id:
        return
    cache = load_cache()
    cache[str(report_id)] = {
        'ts': _now_ts(),
        'status': status or '',
        'raw': raw or {},
        'error': error or ''
    }
    save_cache(cache)


# ============================================================================
# MinistrySafe API client
# ============================================================================
def call_ministrysafe(report_id, force_refresh=False):
    """Fetch the live status for a single report from MS.

    Returns dict:
      { 'status': '<MS status string>',
        'mappedLabel': '<display label>',
        'mappedColor': '<color hex>',
        'raw': {...},
        'error': '' | '<error message>',
        'cached': bool,
        'cachedAge': int_seconds }

    Errors are returned in the dict, not raised — keeps the AJAX caller
    able to render every row even when MS misbehaves on a few of them.
    """
    if not report_id:
        return {
            'status': '',
            'mappedLabel': 'Error Sending - Check TouchPoint',
            'mappedColor': '#dc2626',
            'raw': {},
            'error': 'No ReportID on this record (TouchPoint never received MS confirmation)',
            'cached': False,
            'cachedAge': 0
        }

    if not force_refresh:
        cached = get_cached(report_id)
        if cached:
            mapped = map_status(cached.get('status', ''))
            mapped['raw'] = cached.get('raw', {})
            mapped['error'] = cached.get('error', '')
            mapped['cached'] = True
            mapped['cachedAge'] = _now_ts() - int(cached.get('ts', 0))
            return mapped

    token = get_ms_api_token()
    if not token:
        return {
            'status': '',
            'mappedLabel': 'API token not configured',
            'mappedColor': '#dc2626',
            'raw': {},
            'error': 'Set the MinistrySafe API token on the Settings tab',
            'cached': False,
            'cachedAge': 0
        }

    url = "{0}/background_checks/{1}".format(MS_API_BASE, report_id)
    status = ''
    raw = {}
    error = ''
    try:
        req = urllib2.Request(url)
        req.add_header("Authorization", "Bearer {0}".format(token))
        req.add_header("Content-Type", "application/json")
        req.add_header("User-Agent", "TPxi-BGCheckTracker/{0}".format(APP_VERSION))
        resp = urllib2.urlopen(req, timeout=15)
        code = resp.getcode()
        if code == 200:
            body = resp.read()
            try:
                raw = json.loads(body) or {}
            except Exception:
                raw = {}
            status = str(raw.get('status', 'unknown'))
            put_cached(report_id, status, raw=raw, error='')
        else:
            error = 'HTTP {0}'.format(code)
            put_cached(report_id, '', raw={}, error=error)
    except urllib2.HTTPError as e:
        error = 'HTTP {0}'.format(e.code)
        # 401 / 403 = bad token, no point caching the failure beyond this call
        if e.code not in (401, 403):
            put_cached(report_id, '', raw={}, error=error)
    except urllib2.URLError as e:
        error = 'Network: ' + str(getattr(e, 'reason', e))
    except Exception as e:
        error = 'Unexpected: ' + str(e)

    mapped = map_status(status)
    mapped['raw'] = raw
    mapped['error'] = error
    mapped['cached'] = False
    mapped['cachedAge'] = 0
    return mapped


def resend_ms_invite(report_id):
    """Trigger an applicant-invite resend via MS. The actual endpoint
    name has varied across MS API revisions; we hit the documented
    POST .../background_checks/{id}/resend endpoint and report the HTTP
    response back to the UI verbatim so an admin can troubleshoot."""
    if not report_id:
        return {'success': False, 'message': 'No ReportID for this record'}
    token = get_ms_api_token()
    if not token:
        return {'success': False, 'message': 'API token not configured — set it on the Settings tab'}

    url = "{0}/background_checks/{1}/resend".format(MS_API_BASE, report_id)
    try:
        req = urllib2.Request(url, data='')  # empty POST body
        req.add_header("Authorization", "Bearer {0}".format(token))
        req.add_header("Content-Type", "application/json")
        req.add_header("User-Agent", "TPxi-BGCheckTracker/{0}".format(APP_VERSION))
        req.get_method = lambda: 'POST'
        resp = urllib2.urlopen(req, timeout=15)
        code = resp.getcode()
        if 200 <= code < 300:
            # Invalidate the cache for this report so the next status
            # check pulls fresh data reflecting the resend.
            cache = load_cache()
            if str(report_id) in cache:
                del cache[str(report_id)]
                save_cache(cache)
            return {'success': True, 'message': 'Invite resent (HTTP {0})'.format(code)}
        return {'success': False, 'message': 'HTTP {0}'.format(code)}
    except urllib2.HTTPError as e:
        body = ''
        try:
            body = e.read()
        except Exception:
            pass
        return {'success': False, 'message': 'HTTP {0}: {1}'.format(e.code, body[:200])}
    except urllib2.URLError as e:
        return {'success': False, 'message': 'Network: ' + str(getattr(e, 'reason', e))}
    except Exception as e:
        return {'success': False, 'message': 'Unexpected: ' + str(e)}


def test_api_connection():
    """Light-weight ping: try to list a single background check just to
    see if auth works. Doesn't actually need the response, only the HTTP
    status code."""
    token = get_ms_api_token()
    if not token:
        return {'success': False, 'message': 'No token saved yet.'}
    url = "{0}/background_checks?limit=1".format(MS_API_BASE)
    try:
        req = urllib2.Request(url)
        req.add_header("Authorization", "Bearer {0}".format(token))
        req.add_header("Content-Type", "application/json")
        req.add_header("User-Agent", "TPxi-BGCheckTracker/{0}".format(APP_VERSION))
        resp = urllib2.urlopen(req, timeout=10)
        return {'success': True, 'message': 'Connected to MinistrySafe (HTTP {0})'.format(resp.getcode())}
    except urllib2.HTTPError as e:
        if e.code in (401, 403):
            return {'success': False, 'message': 'Token rejected (HTTP {0}). Double-check the API key on the Settings tab.'.format(e.code)}
        return {'success': False, 'message': 'HTTP {0}'.format(e.code)}
    except urllib2.URLError as e:
        return {'success': False, 'message': 'Network: ' + str(getattr(e, 'reason', e))}
    except Exception as e:
        return {'success': False, 'message': 'Unexpected: ' + str(e)}


# ============================================================================
# Status mapping — single source of truth shared between row render and
# refresh actions so the live + cached statuses always render the same.
# ============================================================================
def map_status(api_status):
    """MS raw status string -> { mappedLabel, mappedColor, mappedKind }.
    mappedKind is 'ok' | 'pending' | 'action' | 'error' | 'unknown' so the
    UI can group / filter without re-deriving from the label text."""
    s = (api_status or '').lower().strip()
    if s == 'complete':
        return {'status': s, 'mappedLabel': 'Awaiting Church Review',
                'mappedColor': '#ea580c', 'mappedKind': 'pending'}
    if s in ('processing', 'pending', 'in_progress'):
        return {'status': s, 'mappedLabel': 'Processing at MinistrySafe',
                'mappedColor': '#2563eb', 'mappedKind': 'pending'}
    if s in ('awaiting_applicant', 'awaiting_invitee', 'started', 'invited'):
        return {'status': s, 'mappedLabel': 'Awaiting Applicant Action',
                'mappedColor': '#7c3aed', 'mappedKind': 'action'}
    if s == 'cancelled':
        return {'status': s, 'mappedLabel': 'Cancelled',
                'mappedColor': '#dc2626', 'mappedKind': 'error'}
    if s == 'expired':
        return {'status': s, 'mappedLabel': 'Expired',
                'mappedColor': '#dc2626', 'mappedKind': 'error'}
    if s == '':
        return {'status': '', 'mappedLabel': 'Not yet checked',
                'mappedColor': '#6b7280', 'mappedKind': 'unknown'}
    # Unknown but recognized; show verbatim so admin can see what MS sent
    return {'status': s, 'mappedLabel': 'MS: ' + s,
            'mappedColor': '#6b7280', 'mappedKind': 'unknown'}


# ============================================================================
# SQL — pending BG list (same as Compliance Dashboard, kept compatible)
# ============================================================================
def pending_bg_sql():
    return '''
        WITH LatestPending AS (
            SELECT
                bc.PeopleId,
                bc.Updated,
                bc.ReportLabelID,
                bc.ReportID,
                ROW_NUMBER() OVER (PARTITION BY bc.PeopleId ORDER BY bc.Updated DESC) AS rn
            FROM BackgroundChecks bc
            WHERE bc.ApprovalStatus = 'Pending' AND bc.ReportTypeId <> 3
        )
        SELECT
            p.PeopleId,
            p.Name,
            p.Name2,
            p.Age,
            ms.Description AS MemberStatus,
            lp.ReportID,
            bcl.Description AS LabelText,
            lp.Updated
        FROM LatestPending lp
        INNER JOIN People p ON p.PeopleId = lp.PeopleId
        LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
        LEFT JOIN lookup.BackgroundCheckLabels bcl ON bcl.Id = lp.ReportLabelID
        WHERE lp.rn = 1
        ORDER BY lp.Updated DESC;
    '''


def load_pending_rows():
    """Run the SQL + attach cached MS status (if present in our blob).
    Live API is NOT called here — that happens on the per-row Refresh /
    Refresh All actions so loading the page stays fast."""
    rows = []
    cache = load_cache()
    for r in q.QuerySql(pending_bg_sql()):
        report_id = str(r.ReportID or '').strip()
        cached = cache.get(report_id) if report_id else None
        if cached:
            mapped = map_status(cached.get('status', ''))
            live = {
                'label': mapped['mappedLabel'],
                'color': mapped['mappedColor'],
                'kind': mapped['mappedKind'],
                'rawStatus': cached.get('status', ''),
                'cached': True,
                'cachedAgeSeconds': _now_ts() - int(cached.get('ts', 0)),
                'error': cached.get('error', '')
            }
        else:
            live = None  # UI shows "(not checked)" with a Refresh button

        # Build the TouchPoint-cached column label. TP only knows "Pending"
        # here (the SQL filtered on ApprovalStatus='Pending'), so we just
        # report the date so admins can see "TP said Pending on X".
        updated_str = ''
        try:
            if r.Updated:
                updated_str = r.Updated.strftime('%Y-%m-%d %H:%M')
        except Exception:
            updated_str = str(r.Updated or '')

        rows.append({
            'peopleId': r.PeopleId,
            'name': r.Name,
            'nameLastFirst': r.Name2 or r.Name,
            'age': r.Age,
            'memberStatus': r.MemberStatus or '',
            'reportId': report_id,
            'label': r.LabelText or '',
            'pendingSince': updated_str,
            'tpStatus': 'Pending',
            'tpStatusColor': '#2563eb',
            'liveStatus': live,
            'msReportUrl': ('https://safetysystem.ministrysafe.com/background_checks/' + report_id) if report_id else '',
            'personUrl': '/Person2/{0}#tab-volunteer'.format(r.PeopleId)
        })
    return rows


# ============================================================================
# AJAX HANDLER
# ============================================================================
if model.HttpMethod == "post":
    action = str(getattr(Data, 'action', '') or '')

    if action == 'load_pending':
        try:
            rows = load_pending_rows()
            print safe_json({
                'success': True,
                'count': len(rows),
                'rows': rows,
                'apiTokenConfigured': bool(get_ms_api_token())
            })
        except Exception as e:
            print safe_json({'success': False, 'message': str(e)})

    elif action == 'refresh_status':
        # Refresh a single row's live MS status. Pass force_refresh=true
        # to bypass the 5-minute cache (used by the "Refresh" button).
        report_id = str(getattr(Data, 'report_id', '') or '').strip()
        force = str(getattr(Data, 'force', 'true')).lower() in ('true', '1', 'yes')
        try:
            result = call_ministrysafe(report_id, force_refresh=force)
            print safe_json({
                'success': True,
                'reportId': report_id,
                'live': {
                    'label': result['mappedLabel'],
                    'color': result['mappedColor'],
                    'kind': result['mappedKind'],
                    'rawStatus': result['status'],
                    'cached': result['cached'],
                    'cachedAgeSeconds': result['cachedAge'],
                    'error': result.get('error', '')
                }
            })
        except Exception as e:
            print safe_json({'success': False, 'message': str(e)})

    elif action == 'refresh_all':
        # Refresh every pending row's live status. Throttled with a small
        # sleep between calls to avoid hammering MS. Returns the full
        # updated row list so the UI can just replace its state.
        try:
            rows = load_pending_rows()
            results = {}
            for i, row in enumerate(rows):
                rid = row.get('reportId')
                if not rid:
                    results[str(row['peopleId'])] = {
                        'label': 'Error Sending - Check TouchPoint',
                        'color': '#dc2626',
                        'kind': 'error',
                        'rawStatus': '',
                        'cached': False,
                        'cachedAgeSeconds': 0,
                        'error': 'No ReportID'
                    }
                    continue
                res = call_ministrysafe(rid, force_refresh=True)
                results[str(row['peopleId'])] = {
                    'label': res['mappedLabel'],
                    'color': res['mappedColor'],
                    'kind': res['mappedKind'],
                    'rawStatus': res['status'],
                    'cached': res['cached'],
                    'cachedAgeSeconds': res['cachedAge'],
                    'error': res.get('error', '')
                }
                # 200ms breathing room between MS calls
                if i < len(rows) - 1:
                    time.sleep(0.2)
            print safe_json({'success': True, 'refreshed': len(results), 'live': results})
        except Exception as e:
            print safe_json({'success': False, 'message': str(e)})

    elif action == 'resend_invite':
        report_id = str(getattr(Data, 'report_id', '') or '').strip()
        try:
            print safe_json(resend_ms_invite(report_id))
        except Exception as e:
            print safe_json({'success': False, 'message': str(e)})

    elif action == 'get_settings':
        # Returns just enough for the Settings tab to know whether a
        # token is set without leaking the token itself.
        try:
            tok = get_ms_api_token()
            masked = ''
            if tok:
                masked = '*' * max(0, len(tok) - 4) + tok[-4:]
            print safe_json({
                'success': True,
                'tokenConfigured': bool(tok),
                'tokenMasked': masked,
                'version': APP_VERSION
            })
        except Exception as e:
            print safe_json({'success': False, 'message': str(e)})

    elif action == 'save_settings':
        # Token comes as form data. Empty string clears the token.
        try:
            tok = str(getattr(Data, 'api_token', '') or '').strip()
            set_ms_api_token(tok)
            print safe_json({'success': True, 'message': 'Saved.'})
        except Exception as e:
            print safe_json({'success': False, 'message': str(e)})

    elif action == 'test_api':
        try:
            print safe_json(test_api_connection())
        except Exception as e:
            print safe_json({'success': False, 'message': str(e)})

    elif action == 'apply_update':
        # Auto-update: fetch the latest source from the workers.dev
        # mirror and overwrite the installed PythonContent slot.
        new_code = ''
        try:
            fetch_url = DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID
            new_code = str(model.RestGet(fetch_url, {}))
        except Exception as fe:
            print safe_json({'success': False, 'message': 'Failed to fetch update: ' + str(fe)})
        else:
            if not new_code or len(new_code) < 200:
                print safe_json({'success': False, 'message': 'Invalid or empty script code received'})
            else:
                target_name = get_script_name() or DC_SCRIPT_ID
                try:
                    model.WriteContentPython(target_name, new_code)
                    print safe_json({'success': True, 'message': 'Updated ' + target_name + '. Reload the page.'})
                except Exception as we:
                    print safe_json({'success': False, 'message': 'Write failed: ' + str(we)})

    else:
        print safe_json({'success': False, 'message': 'Unknown action: ' + action})


# ============================================================================
# MAIN PAGE DISPLAY
# ============================================================================
else:
    html = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>BG Check Tracker</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        .bgt-root { font-family: 'Segoe UI', -apple-system, sans-serif; color: #1f2937; }
        .bgt-root h2 { font-weight: 600; margin: 0 0 4px; }
        .bgt-root .sub { color: #6b7280; margin-bottom: 18px; }
        .bgt-tabs { display: flex; gap: 4px; border-bottom: 2px solid #e5e7eb; margin-bottom: 16px; }
        .bgt-tab { padding: 10px 16px; cursor: pointer; border: none; background: transparent; font-weight: 600; color: #6b7280; border-bottom: 2px solid transparent; margin-bottom: -2px; }
        .bgt-tab.active { color: #0078d4; border-bottom-color: #0078d4; }
        .bgt-tab:hover:not(.active) { color: #1f2937; }
        .bgt-view { display: none; }
        .bgt-view.active { display: block; }

        .bgt-actions { display: flex; gap: 8px; align-items: center; margin-bottom: 12px; flex-wrap: wrap; }
        .bgt-btn { padding: 7px 14px; border-radius: 5px; border: 1px solid #d1d5db; background: #fff; cursor: pointer; font-size: 13px; font-weight: 500; }
        .bgt-btn:hover { background: #f3f4f6; }
        .bgt-btn.primary { background: #0078d4; color: #fff; border-color: #0078d4; }
        .bgt-btn.primary:hover { background: #106ebe; }
        .bgt-btn.warn { background: #fff7ed; color: #9a3412; border-color: #fed7aa; }
        .bgt-btn:disabled { opacity: 0.5; cursor: wait; }
        .bgt-pill { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; color: #fff; white-space: nowrap; }

        .bgt-warn-banner { background: #fff7ed; border: 1px solid #fed7aa; color: #9a3412; padding: 10px 14px; border-radius: 6px; margin-bottom: 14px; font-size: 13px; }

        .bgt-table { width: 100%; border-collapse: collapse; background: #fff; }
        .bgt-table th, .bgt-table td { padding: 8px 10px; text-align: left; font-size: 13px; border-bottom: 1px solid #e5e7eb; vertical-align: middle; }
        .bgt-table th { background: #f9fafb; font-weight: 700; color: #4b5563; font-size: 11px; text-transform: uppercase; letter-spacing: 0.4px; }
        .bgt-table tr:hover { background: #fafbfc; }
        .bgt-table a { color: #0078d4; text-decoration: none; }
        .bgt-table a:hover { text-decoration: underline; }

        .bgt-drift { background: rgba(220,38,38,0.06); }
        .bgt-drift td { border-left: 3px solid #dc2626; }
        .bgt-drift td:first-child { padding-left: 8px; }

        .bgt-row-actions { display: flex; gap: 4px; }
        .bgt-icon-btn { background: none; border: 1px solid #e5e7eb; border-radius: 4px; padding: 3px 7px; cursor: pointer; font-size: 11px; color: #6b7280; }
        .bgt-icon-btn:hover { background: #f3f4f6; color: #1f2937; border-color: #d1d5db; }
        .bgt-icon-btn:disabled { opacity: 0.5; cursor: wait; }
        .bgt-status-cell .raw { display: block; font-size: 10px; color: #9ca3af; margin-top: 2px; }
        .bgt-cached-note { font-size: 10px; color: #9ca3af; }

        .bgt-key-legend { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px; padding: 10px 14px; margin-bottom: 14px; font-size: 12px; }
        .bgt-key-legend strong { display: inline-block; min-width: 200px; }

        .bgt-settings { max-width: 720px; }
        .bgt-settings .field { margin-bottom: 16px; }
        .bgt-settings label { display: block; font-weight: 600; margin-bottom: 4px; color: #1f2937; }
        .bgt-settings input[type=text], .bgt-settings input[type=password] { width: 100%; padding: 8px 10px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 13px; font-family: monospace; }
        .bgt-settings .help { background: #eff6ff; border: 1px solid #bfdbfe; border-left: 4px solid #2563eb; padding: 12px 16px; border-radius: 6px; font-size: 13px; line-height: 1.5; margin-top: 18px; }
        .bgt-settings .help h4 { margin: 0 0 8px; color: #1e3a8a; font-size: 14px; }
        .bgt-settings .help ol { margin: 0; padding-left: 22px; }
        .bgt-settings .help li { margin-bottom: 4px; }
        .bgt-settings .help code { background: #fff; padding: 1px 6px; border-radius: 3px; font-size: 12px; border: 1px solid #dbeafe; }

        .bgt-empty { text-align: center; padding: 40px 20px; color: #9ca3af; }
        .bgt-loading { text-align: center; padding: 30px; color: #6b7280; }
        .bgt-spinner { display: inline-block; width: 20px; height: 20px; border: 3px solid #e5e7eb; border-top-color: #0078d4; border-radius: 50%; animation: bgt-spin 0.8s linear infinite; vertical-align: middle; margin-right: 8px; }
        @keyframes bgt-spin { to { transform: rotate(360deg); } }

        .bgt-toast { position: fixed; bottom: 24px; right: 24px; background: #1f2937; color: #fff; padding: 10px 16px; border-radius: 6px; font-size: 13px; box-shadow: 0 4px 16px rgba(0,0,0,0.2); z-index: 10000; opacity: 0; transition: opacity 0.2s; }
        .bgt-toast.show { opacity: 1; }
        .bgt-toast.error { background: #dc2626; }
        .bgt-toast.success { background: #059669; }
    </style>
</head>
<body>

<div class="bgt-root">

    <!-- Auto-update banner -->
    <div id="appUpdateBanner" style="display:none;background:#e8f0fd;border:1px solid #cfe3ff;border-radius:8px;padding:10px 14px;margin-bottom:12px;align-items:center;gap:10px;"></div>

    <h2><i class="bi bi-shield-check"></i> Background Check Tracker
        <span style="font-size:13px; font-weight:normal; color:#6b7280;">v__APP_VERSION__</span>
    </h2>
    <p class="sub">Triage in-flight checks. Pulls TouchPoint's cached status AND the live MinistrySafe status side-by-side so you can spot sync drift fast.</p>

    <div class="bgt-tabs">
        <button class="bgt-tab active" data-tab="pending">Pending BG Checks</button>
        <button class="bgt-tab" data-tab="settings">Settings</button>
    </div>

    <!-- =================== PENDING BG VIEW =================== -->
    <div class="bgt-view active" id="view-pending">

        <div class="bgt-key-legend">
            <strong>Awaiting Church Review</strong> <span class="bgt-pill" style="background:#ea580c;">complete</span> &nbsp; MS finished. Action needed from church staff<br>
            <strong>Processing at MinistrySafe</strong> <span class="bgt-pill" style="background:#2563eb;">processing</span> &nbsp; MS still working<br>
            <strong>Awaiting Applicant Action</strong> <span class="bgt-pill" style="background:#7c3aed;">awaiting_applicant</span> &nbsp; Applicant hasn't completed their portion<br>
            <strong>Cancelled / Expired</strong> <span class="bgt-pill" style="background:#dc2626;">cancelled</span> &nbsp; Resend invite or restart<br>
            <strong>Highlighted rows</strong> <span class="bgt-pill" style="background:#dc2626;">drift</span> &nbsp; TouchPoint says Pending but MS says complete / failed
        </div>

        <div id="apiNotConfigured" class="bgt-warn-banner" style="display:none;">
            <i class="bi bi-exclamation-triangle"></i>
            <strong>MinistrySafe API token not configured.</strong>
            Live status lookups won't work. Switch to the Settings tab to add your token.
        </div>

        <div class="bgt-actions">
            <button class="bgt-btn primary" id="btnRefreshAll"><i class="bi bi-arrow-repeat"></i> Refresh All Live Status</button>
            <button class="bgt-btn" id="btnReload"><i class="bi bi-arrow-clockwise"></i> Reload List</button>
            <span style="margin-left:auto; font-size:12px; color:#6b7280;" id="rowCount"></span>
        </div>

        <div id="pendingTableWrap"></div>
    </div>

    <!-- =================== SETTINGS VIEW =================== -->
    <div class="bgt-view" id="view-settings">
        <div class="bgt-settings">
            <h3 style="margin-top:0;">MinistrySafe API Settings</h3>

            <!-- Wrapping the password field in a <form> silences Chrome's
                 [DOM] "password field is not contained in a form" warning.
                 onsubmit returns false so Enter triggers Save instead of
                 a real GET that would unload the page. -->
            <form onsubmit="document.getElementById('btnSaveSettings').click(); return false;" autocomplete="off">
                <div class="field">
                    <label for="apiToken">API Token</label>
                    <input type="password" id="apiToken" name="apiToken" placeholder="Paste your MinistrySafe API token here" autocomplete="new-password">
                    <div id="tokenStatus" style="margin-top:6px; font-size:12px; color:#6b7280;"></div>
                </div>

                <div class="bgt-actions">
                    <button type="button" class="bgt-btn primary" id="btnSaveSettings"><i class="bi bi-save"></i> Save Token</button>
                    <button type="button" class="bgt-btn" id="btnTestApi"><i class="bi bi-plug"></i> Test Connection</button>
                    <button type="button" class="bgt-btn warn" id="btnClearToken"><i class="bi bi-trash"></i> Clear Saved Token</button>
                </div>
            </form>

            <div class="help">
                <h4>How to get your MinistrySafe API token</h4>
                <ol>
                    <li>Sign in to <a href="https://safetysystem.ministrysafe.com" target="_blank">safetysystem.ministrysafe.com</a> as an admin user.</li>
                    <li>Click your name (top right) and choose <code>Account Settings</code>.</li>
                    <li>Scroll to <code>API Access</code> (lower section of the page).</li>
                    <li>If a token already exists, click <code>Reveal</code> and copy it. Otherwise click <code>Generate API Token</code> first.</li>
                    <li>Paste the token into the field above and click <strong>Save Token</strong>.</li>
                    <li>Click <strong>Test Connection</strong> to confirm the token works. A green toast confirms success.</li>
                </ol>
                <p style="margin: 8px 0 0;"><strong>Don't see API Access?</strong> Your MinistrySafe account may not have API permission enabled. Email <a href="mailto:support@ministrysafe.com">support@ministrysafe.com</a> to request it — they enable it on most paid plans for free.</p>
                <p style="margin: 8px 0 0;"><strong>Rotating the token?</strong> Generate a new one in MinistrySafe FIRST, paste it here, save, and verify Test Connection succeeds. The old token in this app will continue to work until you replace it.</p>
            </div>

            <div class="help" style="background:#fff7ed; border-color:#fed7aa; border-left-color:#ea580c;">
                <h4 style="color:#9a3412;">What does this token let the script do?</h4>
                <ul style="margin: 0; padding-left: 22px;">
                    <li>READ background check status (live, per applicant)</li>
                    <li>RESEND the applicant invite (when status is "Awaiting Applicant Action")</li>
                </ul>
                <p style="margin: 8px 0 0;">The token is stored in TouchPoint's <code>Setting</code> table under <code>__SETTING_NAME__</code>. It is NOT exposed in any returned page, log line, or AJAX response.</p>
            </div>
        </div>
    </div>

</div>

<div id="bgtToast" class="bgt-toast"></div>

<script>
// ============================================================================
// State + URL helpers
// ============================================================================
var APP_VERSION = '__APP_VERSION__';
var DC_SCRIPT_ID = '__DC_SCRIPT_ID__';
var DC_API_BASE = '__DC_API_BASE__';
var SCRIPT_NAME = (function() {
    try {
        var parts = (window.location.pathname || '').split('/');
        for (var i = 0; i < parts.length; i++) {
            if ((parts[i] === 'PyScript' || parts[i] === 'PyScriptForm') && i + 1 < parts.length && parts[i + 1]) {
                return parts[i + 1].split('?')[0].split('#')[0];
            }
        }
    } catch(e) {}
    return DC_SCRIPT_ID;
})();
var APP_UPDATE_AVAILABLE = false;
var APP_LATEST_VERSION = '';
var scriptUrl = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
var state = { rows: [], apiTokenConfigured: false };

function ajax(action, params, cb) {
    params = params || {};
    params.action = action;
    params.script_name = SCRIPT_NAME;
    var body = Object.keys(params).map(function(k) {
        return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    }).join('&');
    var xhr = new XMLHttpRequest();
    xhr.open('POST', scriptUrl, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        try {
            var data = JSON.parse(xhr.responseText);
            if (cb) cb(data);
        } catch(e) {
            console.error('[BGT] Parse error', e, xhr.responseText.substring(0, 400));
            toast('Error parsing response (see console)', 'error');
        }
    };
    xhr.send(body);
}

function toast(msg, type) {
    var el = document.getElementById('bgtToast');
    el.textContent = msg;
    el.className = 'bgt-toast show ' + (type || '');
    setTimeout(function() { el.classList.remove('show'); }, 3000);
}

function escapeHtml(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function fmtAgeSeconds(sec) {
    if (!sec || sec < 0) return '';
    if (sec < 60) return sec + 's ago';
    if (sec < 3600) return Math.floor(sec / 60) + 'm ago';
    return Math.floor(sec / 3600) + 'h ago';
}

// ============================================================================
// Tabs
// ============================================================================
document.querySelectorAll('.bgt-tab').forEach(function(btn) {
    btn.addEventListener('click', function() {
        document.querySelectorAll('.bgt-tab').forEach(function(t) { t.classList.remove('active'); });
        document.querySelectorAll('.bgt-view').forEach(function(v) { v.classList.remove('active'); });
        btn.classList.add('active');
        var name = btn.dataset.tab;
        document.getElementById('view-' + name).classList.add('active');
        if (name === 'settings') loadSettings();
    });
});

// ============================================================================
// Pending BG list
// ============================================================================
function loadPending() {
    document.getElementById('pendingTableWrap').innerHTML =
        '<div class="bgt-loading"><span class="bgt-spinner"></span>Loading pending background checks...</div>';
    ajax('load_pending', {}, function(resp) {
        if (!resp.success) {
            document.getElementById('pendingTableWrap').innerHTML =
                '<div class="bgt-empty">Failed to load: ' + escapeHtml(resp.message || 'unknown error') + '</div>';
            return;
        }
        state.rows = resp.rows || [];
        state.apiTokenConfigured = !!resp.apiTokenConfigured;
        document.getElementById('apiNotConfigured').style.display = state.apiTokenConfigured ? 'none' : 'block';
        renderPendingTable();

        // Auto-fetch live MS status for any row that doesn't have a
        // fresh cached value. Done after the page paints so the table
        // is interactive immediately and the live column fills in as
        // the batch completes. Skipped when no token is configured
        // (would just produce a wall of errors) or when every row
        // already has cached data (cache TTL still in window).
        if (state.apiTokenConfigured && state.rows.length > 0) {
            var needsLive = state.rows.some(function(r) { return !r.liveStatus; });
            if (needsLive) setTimeout(autoRefreshLive, 150);
        }
    });
}

// Background variant of the Refresh All click — same server call, but
// the button shows a subtle inline progress state instead of taking
// over the row count area, since the user didn't explicitly ask for it.
function autoRefreshLive() {
    var btn = document.getElementById('btnRefreshAll');
    if (!btn || btn.disabled) return;
    btn.disabled = true;
    btn.innerHTML = '<span class="bgt-spinner"></span>Fetching live status for ' + state.rows.length + ' rows...';
    ajax('refresh_all', {}, function(resp) {
        btn.disabled = false;
        btn.innerHTML = '<i class="bi bi-arrow-repeat"></i> Refresh All Live Status';
        if (!resp.success) return;
        var live = resp.live || {};
        for (var i = 0; i < state.rows.length; i++) {
            var pid = state.rows[i].peopleId;
            if (live[String(pid)]) state.rows[i].liveStatus = live[String(pid)];
        }
        renderPendingTable();
    });
}

function renderPendingTable() {
    var rows = state.rows;
    document.getElementById('rowCount').textContent = rows.length + ' pending';
    if (!rows.length) {
        document.getElementById('pendingTableWrap').innerHTML =
            '<div class="bgt-empty"><i class="bi bi-check-circle" style="font-size:32px; color:#10b981;"></i><br><br>No pending background checks. Nice!</div>';
        return;
    }
    var html = '<table class="bgt-table"><thead><tr>';
    html += '<th>Name</th>';
    html += '<th>Church Membership</th>';
    html += '<th>TP Cached</th>';
    html += '<th>Live MS Status</th>';
    html += '<th>Label</th>';
    html += '<th>Pending Since</th>';
    html += '<th>Actions</th>';
    html += '</tr></thead><tbody>';
    for (var i = 0; i < rows.length; i++) {
        html += renderRow(rows[i]);
    }
    html += '</tbody></table>';
    document.getElementById('pendingTableWrap').innerHTML = html;
}

function renderRow(row) {
    var live = row.liveStatus;
    var driftClass = '';
    // Drift detection: TP says Pending, MS says complete / cancelled / expired
    if (live && (live.kind === 'pending' || live.kind === 'error') && live.rawStatus === 'complete') driftClass = 'bgt-drift';
    if (live && live.kind === 'error') driftClass = 'bgt-drift';

    var html = '<tr class="' + driftClass + '" data-pid="' + row.peopleId + '" data-report="' + escapeHtml(row.reportId) + '">';
    html += '<td><a href="' + row.personUrl + '" target="_blank">' + escapeHtml(row.name) + '</a>';
    if (row.age) html += ' <span style="color:#9ca3af; font-size:11px;">(' + row.age + ')</span>';
    html += '</td>';
    html += '<td>' + escapeHtml(row.memberStatus) + '</td>';
    html += '<td><span class="bgt-pill" style="background:' + row.tpStatusColor + ';">Pending</span></td>';
    html += '<td class="bgt-status-cell">' + renderLiveCell(live) + '</td>';
    html += '<td>' + escapeHtml(row.label) + '</td>';
    html += '<td>' + escapeHtml(row.pendingSince) + '</td>';
    html += '<td><div class="bgt-row-actions">';
    html += '<button class="bgt-icon-btn" data-act="refresh" title="Refresh live status"><i class="bi bi-arrow-clockwise"></i></button>';
    if (live && live.kind === 'action') {
        html += '<button class="bgt-icon-btn" data-act="resend" title="Resend invite to applicant"><i class="bi bi-envelope-arrow-up"></i></button>';
    }
    if (row.msReportUrl) {
        html += '<a class="bgt-icon-btn" href="' + row.msReportUrl + '" target="_blank" title="Open in MinistrySafe"><i class="bi bi-box-arrow-up-right"></i></a>';
    }
    html += '</div></td>';
    html += '</tr>';
    return html;
}

function renderLiveCell(live) {
    if (!live) {
        return '<span style="color:#9ca3af; font-style:italic;">Not checked</span>';
    }
    var html = '<span class="bgt-pill" style="background:' + live.color + ';">' + escapeHtml(live.label) + '</span>';
    if (live.rawStatus && live.kind !== 'unknown' && live.label.indexOf(live.rawStatus) === -1) {
        html += '<span class="raw">MS raw: ' + escapeHtml(live.rawStatus) + '</span>';
    }
    if (live.cached) {
        html += '<span class="bgt-cached-note">cached ' + fmtAgeSeconds(live.cachedAgeSeconds) + '</span>';
    }
    if (live.error) {
        html += '<span class="raw" style="color:#dc2626;">' + escapeHtml(live.error) + '</span>';
    }
    return html;
}

// Delegated row-action handler
document.getElementById('pendingTableWrap').addEventListener('click', function(e) {
    var btn = e.target.closest('[data-act]');
    if (!btn) return;
    var tr = btn.closest('tr');
    var pid = parseInt(tr.dataset.pid, 10);
    var rid = tr.dataset.report;
    var act = btn.dataset.act;
    if (act === 'refresh') refreshRow(pid, rid, btn);
    else if (act === 'resend') resendInvite(pid, rid, btn);
});

function refreshRow(peopleId, reportId, btn) {
    btn.disabled = true;
    btn.innerHTML = '<i class="bi bi-hourglass"></i>';
    ajax('refresh_status', { report_id: reportId, force: 'true' }, function(resp) {
        btn.disabled = false;
        btn.innerHTML = '<i class="bi bi-arrow-clockwise"></i>';
        if (!resp.success) {
            toast('Refresh failed: ' + (resp.message || ''), 'error');
            return;
        }
        // Find the row in state + update it
        for (var i = 0; i < state.rows.length; i++) {
            if (state.rows[i].peopleId === peopleId) {
                state.rows[i].liveStatus = resp.live;
                break;
            }
        }
        renderPendingTable();
        toast('Refreshed live status', 'success');
    });
}

function resendInvite(peopleId, reportId, btn) {
    if (!confirm('Resend the MinistrySafe invite for this applicant?')) return;
    btn.disabled = true;
    btn.innerHTML = '<i class="bi bi-hourglass"></i>';
    ajax('resend_invite', { report_id: reportId }, function(resp) {
        btn.disabled = false;
        btn.innerHTML = '<i class="bi bi-envelope-arrow-up"></i>';
        toast(resp.message || (resp.success ? 'Invite resent' : 'Resend failed'), resp.success ? 'success' : 'error');
        if (resp.success) {
            // Force a refresh so the applicant-action label updates if it changes
            refreshRow(peopleId, reportId, btn);
        }
    });
}

document.getElementById('btnReload').addEventListener('click', loadPending);

document.getElementById('btnRefreshAll').addEventListener('click', function() {
    if (!state.rows.length) { toast('Nothing to refresh'); return; }
    if (!state.apiTokenConfigured) {
        toast('Add your API token on the Settings tab first', 'error');
        return;
    }
    var btn = this;
    btn.disabled = true;
    btn.innerHTML = '<span class="bgt-spinner"></span>Refreshing ' + state.rows.length + ' rows...';
    ajax('refresh_all', {}, function(resp) {
        btn.disabled = false;
        btn.innerHTML = '<i class="bi bi-arrow-repeat"></i> Refresh All Live Status';
        if (!resp.success) {
            toast('Refresh failed: ' + (resp.message || ''), 'error');
            return;
        }
        var live = resp.live || {};
        for (var i = 0; i < state.rows.length; i++) {
            var pid = state.rows[i].peopleId;
            if (live[String(pid)]) state.rows[i].liveStatus = live[String(pid)];
        }
        renderPendingTable();
        toast('Refreshed ' + (resp.refreshed || 0) + ' rows', 'success');
    });
});

// ============================================================================
// Settings
// ============================================================================
function loadSettings() {
    ajax('get_settings', {}, function(resp) {
        if (!resp.success) {
            document.getElementById('tokenStatus').textContent = 'Failed to load settings';
            return;
        }
        var el = document.getElementById('tokenStatus');
        if (resp.tokenConfigured) {
            el.innerHTML = '<i class="bi bi-check-circle-fill" style="color:#10b981;"></i> Token saved (ending in <code>' + escapeHtml(resp.tokenMasked) + '</code>). Leave field blank to keep current value.';
            document.getElementById('apiToken').placeholder = 'Leave blank to keep current token, or paste new to replace';
        } else {
            el.innerHTML = '<i class="bi bi-exclamation-triangle-fill" style="color:#ea580c;"></i> No token saved yet.';
        }
    });
}

document.getElementById('btnSaveSettings').addEventListener('click', function() {
    var tok = document.getElementById('apiToken').value.trim();
    if (!tok) {
        toast('Enter a token first (or use Clear to remove the saved one)', 'error');
        return;
    }
    ajax('save_settings', { api_token: tok }, function(resp) {
        if (resp.success) {
            toast('Token saved', 'success');
            document.getElementById('apiToken').value = '';
            // Update the in-memory flag so the Pending tab banner hides
            // immediately on return. Without this the banner would stay
            // visible until the next load_pending fetch on page reload.
            state.apiTokenConfigured = true;
            document.getElementById('apiNotConfigured').style.display = 'none';
            loadSettings();
        } else {
            toast('Save failed: ' + (resp.message || ''), 'error');
        }
    });
});

document.getElementById('btnTestApi').addEventListener('click', function() {
    // If the user typed a new token but hasn't saved yet, save it first
    // so the test reflects what they're about to commit.
    var tok = document.getElementById('apiToken').value.trim();
    var doTest = function() {
        ajax('test_api', {}, function(resp) {
            toast(resp.message, resp.success ? 'success' : 'error');
        });
    };
    if (tok) {
        ajax('save_settings', { api_token: tok }, function(resp) {
            if (resp.success) {
                document.getElementById('apiToken').value = '';
                loadSettings();
                doTest();
            } else {
                toast('Save failed: ' + (resp.message || ''), 'error');
            }
        });
    } else {
        doTest();
    }
});

document.getElementById('btnClearToken').addEventListener('click', function() {
    if (!confirm('Remove the saved MinistrySafe API token? Live status lookups will stop working until you add a new one.')) return;
    ajax('save_settings', { api_token: '' }, function(resp) {
        if (resp.success) {
            toast('Token cleared', 'success');
            state.apiTokenConfigured = false;
            document.getElementById('apiNotConfigured').style.display = 'block';
            loadSettings();
        } else {
            toast('Clear failed: ' + (resp.message || ''), 'error');
        }
    });
});

// ============================================================================
// Auto-Update (see TPxi/AutoUpdate/README.md)
// ============================================================================
function checkForAppUpdate() {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', DC_API_BASE + '/script-versions', true);
        xhr.timeout = 5000;
        xhr.onreadystatechange = function() {
            if (xhr.readyState !== 4 || xhr.status !== 200) return;
            try {
                var versions = JSON.parse(xhr.responseText);
                var latest = versions[DC_SCRIPT_ID];
                if (latest && latest !== APP_VERSION) {
                    APP_UPDATE_AVAILABLE = true;
                    APP_LATEST_VERSION = latest;
                    renderAppUpdateBanner();
                }
            } catch(e) {}
        };
        xhr.send();
    } catch(e) {}
}

function renderAppUpdateBanner() {
    var b = document.getElementById('appUpdateBanner');
    if (!b || !APP_UPDATE_AVAILABLE) return;
    var h = '<div style="font-size:18px">&#128640;</div>';
    h += '<div style="flex:1;font-size:12px;color:#0078d4">';
    h += '<strong>Update available</strong>';
    h += ' &mdash; you have <code>v' + APP_VERSION + '</code>, latest is <code>v' + APP_LATEST_VERSION + '</code>.';
    h += '</div>';
    h += '<button id="appUpdateBtn" onclick="applyAppUpdate()" style="white-space:nowrap;padding:6px 14px;background:#0078d4;color:#fff;border:0;border-radius:4px;cursor:pointer;">Update Now</button>';
    b.innerHTML = h;
    b.style.display = 'flex';
}

window.applyAppUpdate = function() {
    if (!confirm('Update from v' + APP_VERSION + ' to v' + APP_LATEST_VERSION + '?\\n\\nYour saved settings and cache are preserved.')) return;
    var btn = document.getElementById('appUpdateBtn');
    if (btn) { btn.disabled = true; btn.innerHTML = 'Updating...'; }
    ajax('apply_update', {}, function(r) {
        if (r.success) {
            alert(r.message || 'Updated! Reloading...');
            window.location.reload(true);
        } else {
            alert('Update failed: ' + (r.message || 'Unknown error'));
            if (btn) { btn.disabled = false; btn.innerHTML = 'Update Now'; }
        }
    });
};

// ============================================================================
// Boot
// ============================================================================
checkForAppUpdate();
loadPending();

</script>

</body>
</html>
'''

    # Substitute server-side constants into the HTML template (placeholder
    # pattern avoids the JS-quote-escaping headache of inline format()s).
    html = (html
        .replace('__APP_VERSION__',  APP_VERSION)
        .replace('__DC_SCRIPT_ID__', DC_SCRIPT_ID)
        .replace('__DC_API_BASE__',  DC_API_BASE)
        .replace('__SETTING_NAME__', SETTING_API_TOKEN))

    print html
    model.Form = html
