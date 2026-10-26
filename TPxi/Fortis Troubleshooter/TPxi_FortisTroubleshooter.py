"""
Fortis Troubleshooter
=====================
A self-service diagnostics page for a TouchPoint(R) site running payments
through the Fortis gateway. Three tabs, answering three different questions.

DENIALS  -- why did this payment fail?
  Reads dbo.FortisTransactionLog and turns raw gateway rejections into
  something you can hand to your processor: the category, the offending
  field and its length, the last four of the card, and the verbatim error.
  Nobody has to run SQL to answer "what happened to this gift?".

HEALTH & FLOW  -- is anything getting through at all?
  Reads dbo.FortisWebhookTransactions, the actual stream of transactions
  the gateway reports back. Shows when each payment method last produced a
  transaction, compares the last seven days against the same calendar days
  in previous months, tracks ACH through settlement, and can email you when
  a payment method goes silent.

BLOCKED BEFORE FORTIS  -- did TouchPoint stop it before Fortis saw it?
  TouchPoint runs its own checks (captcha, address, email domain, IP fraud
  score, failed-attempt lockout) before sending anything to Fortis. A payment
  stopped there leaves nothing in either Fortis table and TouchPoint does not
  log the block. This tab shows each check's current setting, the IP scores
  TouchPoint saved and who was signed in on those IPs, and a person lookup
  that says whether signing in would get them past the IP check.
  Typical trigger: "Giving is not supported from your current location."

WHY THE SECOND TAB EXISTS
-------------------------
dbo.FortisTransactionLog is an ERROR-ONLY log: every row is a request the
gateway rejected. That is the right tool for "this person's card declined",
and completely blind to the more expensive failure -- a payment method that
stops being submitted at all. Nothing is attempted, so nothing errors, so an
error log shows a reassuring zero.

This was written after ACH submissions stopped for eight days without a
single logged error. Donors thought they had given; the money never moved;
the denial page showed nothing wrong because nothing had gone wrong -- it
had just stopped happening. The health tab makes silence visible.

READING THE DENIAL DATA
-----------------------
Denials fall into two buckets depending on which endpoint was hit, and the
difference matters when you send examples to your processor:

  * Transaction errors come from the transaction endpoints
    (/v1/transactions/cc/sale/keyed and the ACH equivalent). Those requests
    carry the masked account number, so the LAST 4 IS AVAILABLE.

  * Contact-sync errors come from /v1/contacts. That request carries only
    name and address -- no card at all -- so there is NO last 4 to report.
    The card never entered the picture; the contact record was rejected
    before any transaction existed.

Because account numbers are stored already masked, this page never exposes
a full PAN. There is no PCI concern in sharing its output.

WHAT YOU WILL NEED TO SET
-------------------------
Everything configurable sits in the CONFIGURATION blocks below. At minimum
check ALLOWED_ROLES, CONTACT_API_PREFIX and UTC_OFFSET_HOURS. To get email
alerts, set HEALTH_ALERT_PEOPLE_ID and add the ScheduledTasks block shown
at the bottom of the Health tab.

--Upload Instructions Start--
1. Admin ~ Advanced ~ Special Content ~ Python
2. New Python Script File
3. Name it "FortisTroubleshooter" and paste this code
4. Access via /PyScriptForm/FortisTroubleshooter
   (optionally add to a menu / Blue Toolbar)
--Upload Instructions End--

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

Version: 2.2
"""

import json
import re
import datetime

#####################################################################
#### CONFIGURATION
#####################################################################

PAGE_TITLE = "Fortis Denial Troubleshooter"

# Roles allowed to view this page. User needs at least one of these.
ALLOWED_ROLES = ["Finance", "Admin", "ManageApplication"]

#####################################################################
#### HEALTH TAB CONFIGURATION
#####################################################################
# The denial tab reads FortisTransactionLog, which is ERROR-ONLY. That answers
# "what failed", but it cannot answer "did anything happen at all" -- an
# absence leaves no error to find. In September 2026 ACH submissions stopped
# entirely for eight days and produced no errors, because nothing was ever
# sent. Nobody noticed.
#
# The health tab reads dbo.FortisWebhookTransactions, the actual transaction
# stream Fortis reports back, so it can see silence.

# Displayed timestamps are shifted by this many hours from UTC, which is what
# the webhook epochs are in. Ages and staleness are always computed in UTC, so
# a wrong value here only affects how times READ, never whether we alert.
# Central is -5 during daylight saving and -6 outside it.
UTC_OFFSET_HOURS = -5

# Trailing days compared against the same calendar window in prior months.
# Seven is deliberate: it evens out weekday and weekend patterns.
HEALTH_WINDOW_DAYS = 7

# How many previous months to use as the baseline for that window.
HEALTH_BASELINE_MONTHS = 3

# A method is flagged only when it has produced NOTHING in the trailing
# window AND the baseline says it normally produces at least this many. A
# plain "days since last transaction" rule does not work here: ACH is sparse
# mid-month and has legitimate 12 to 15 day gaps at month boundaries, so it
# would either cry wolf constantly or miss a real outage.
HEALTH_MIN_EXPECTED = 2

# Who to email when a payment method goes silent. 0 disables email; the tab
# still shows the status.
HEALTH_ALERT_PEOPLE_ID = 0
HEALTH_ALERT_FROM_EMAIL = "noreply@yourchurch.org"
HEALTH_ALERT_FROM_NAME = "Fortis Health Check"

# One alert per silent method per this many hours.
HEALTH_ALERT_COOLDOWN_HOURS = 24

HEALTH_STATE_KEY = "FortisHealth_State"

# ACH settlement lifecycle: CREATE(131) -> UPDATE(132) -> UPDATE(134 + SettleDate).
# Anything below 134 after this many days has not settled and is worth chasing.
ACH_SETTLE_WARN_DAYS = 5

# Card authorises and captures in one step and never emits a settlement
# webhook, so its rows legitimately have no SettleDate. Only ACH is tracked
# through the pipeline.
SETTLED_STATUS = 134

# Fortis stores each synced contact with a contact_api_id built from a site
# prefix plus the TouchPoint PeopleId, e.g. "MYCHURCH12345". TouchPoint sets
# the prefix per site, so CHANGE THIS to match yours or contact-sync denials
# will not resolve back to a person.
#
# To find it: open the Denials tab, look at a Street Length or Last Name
# denial, and read the contact_api_id out of the request body.
CONTACT_API_PREFIX = "MYCHURCH"

DEFAULT_DAYS = 90
MAX_DETAIL_ROWS = 1000  # safety cap on the detail table

# Category identifiers (also used as the summary ordering)
CATEGORIES = [
    ("card",     "Card Number Invalid"),
    ("street",   "Street Length/Format"),
    ("phone",    "Phone Format"),
    ("routing",  "Routing Number Invalid (ACH)"),
    ("lastname", "Last Name Empty"),
    ("datavalidation", "Data Validation Failed (generic 422)"),
    ("other",    "Other / Uncategorized"),
]
CATEGORY_LABELS = dict(CATEGORIES)

# Which transaction stream each denial category should be measured AGAINST.
# A raw denial count rises purely with volume, so it cannot tell a worsening
# problem from a busy month. Dividing by the matching transaction count does.
#
# "street" and "lastname" come from the CONTACT SYNC endpoint (/v1/contacts),
# which is not a transaction at all, so there is no honest denominator and
# those cards stay as plain counts rather than inventing a rate.
CATEGORY_DENOMINATOR = {
    "card": "cc",
    "phone": "cc",
    "routing": "ach",
    "street": None,
    "lastname": None,
    "datavalidation": None,
    "other": None,
}


#####################################################################
#### PERMISSION CHECK
#####################################################################

def user_allowed():
    try:
        for r in ALLOWED_ROLES:
            if model.UserIsInRole(r):
                return True
    except:
        pass
    return False


#####################################################################
#### PARSING HELPERS
#####################################################################

def _slice_between(text, start_marker, end_marker):
    """Return the substring after start_marker up to (but not including)
    end_marker. If end_marker is not found, return to end of string."""
    if not text:
        return ""
    i = text.find(start_marker)
    if i < 0:
        return ""
    rest = text[i + len(start_marker):]
    if end_marker:
        j = rest.find(end_marker)
        if j >= 0:
            rest = rest[:j]
    return rest.strip()


def extract_endpoint(request_text):
    """Pull the QueryUrl out of the logged Request and return a friendly type."""
    url = _slice_between(request_text, "QueryUrl = ", ",  QueryParameters")
    low = url.lower()
    if "cc/sale" in low:
        friendly = "CC Sale"
    elif "cc/refund" in low:
        friendly = "CC Refund"
    elif "ach/debit" in low:
        friendly = "ACH Debit"
    elif "ach/credit" in low:
        friendly = "ACH Credit"
    elif "/contacts" in low:
        friendly = "Contact Sync"
    elif "/tokens" in low:
        friendly = "Token"
    elif "/transactions/" in low:
        friendly = "Transaction (other)"
    else:
        friendly = "Other"
    return friendly, url


def extract_body_json(request_text):
    """Extract and parse the JSON Body out of the logged Request.
    Returns (dict_or_None, raw_body_string)."""
    body = _slice_between(request_text, "Body = ", ",  Username =")
    if not body:
        return None, ""
    try:
        return json.loads(body), body
    except:
        return None, body


GENERIC_MSG = "data validation failed."


def clean_message(resp):
    """Normalize the Response field into just the validation message.
    'Validation Error: Account Number is not valid - Unprocessable Entity, Http code: 422'
      -> 'Account Number is not valid'
    """
    if not resp:
        return ""
    m = resp.strip()
    if m.lower().startswith("validation error:"):
        m = m[len("validation error:"):].strip()
    # strip trailing ' - <Title>, Http code: NNN'
    m = re.sub(r'\s*-\s*[^,]*,\s*Http code:\s*\d+\s*$', '', m)
    return m.strip()


def get_error_detail(row):
    """Return the most specific human-readable validation message.
    IMPORTANT: FortisError.detail is frequently the generic 'Data Validation
    Failed.' while the Response field carries the real reason (e.g. 'Account
    Number is not valid'), so we prefer Response."""
    cleaned = clean_message(getattr(row, "Response", None))
    if cleaned and cleaned.lower() != GENERIC_MSG:
        return cleaned
    fe = getattr(row, "FortisError", None)
    if fe:
        try:
            j = json.loads(fe)
            d = (j.get("detail") or "").strip()
            if d and d.lower() != GENERIC_MSG:
                return d
        except:
            pass
    if cleaned:
        return cleaned
    return (getattr(row, "ErrorMessage", None) or "").strip()


def categorize_row(row):
    """Categorize using ALL available text (Response + FortisError + ErrorMessage)
    so a generic FortisError.detail never hides the specific Response message."""
    parts = [getattr(row, "Response", "") or "",
             getattr(row, "FortisError", "") or "",
             getattr(row, "ErrorMessage", "") or ""]
    text = " ".join(parts).lower()
    if "routing" in text:
        return "routing"
    if "phone" in text:
        return "phone"
    if "street" in text:
        return "street"
    if "last_name" in text or "last name" in text:
        return "lastname"
    if ("account number is not valid" in text or "unable to determine" in text
            or "card type" in text or "card number" in text or "invalid card" in text):
        return "card"
    if "data validation failed" in text:
        return "datavalidation"
    return "other"


def get_http_status(row):
    """Best-effort HTTP status code from the ErrorMessage / FortisError."""
    fe = getattr(row, "FortisError", None)
    if fe:
        m = re.search(r'"statusCode"\s*:\s*"?(\d{3})', fe)
        if m:
            return m.group(1)
    em = getattr(row, "ErrorMessage", None) or ""
    m = re.search(r'\b(4\d\d|5\d\d)\b', em)
    return m.group(1) if m else ""


def last4_from_account(acct):
    """Given a masked account_number like '************0093' or '***5000',
    return just the trailing 4 visible digits, or '' if none."""
    if not acct:
        return ""
    digits = re.sub(r"\D", "", acct)
    if not digits:
        return ""
    return digits[-4:]


def people_id_from_contact_api(body):
    """Contacts carry contact_api_id '<CONTACT_API_PREFIX><PeopleId>'.

    Returns the PeopleId, or empty if the prefix does not match -- which
    usually means CONTACT_API_PREFIX is still set to the default.
    """
    if not body:
        return None
    cid = body.get("contact_api_id")
    if not cid:
        return None
    m = re.search(r"(\d+)\s*$", str(cid))
    return m.group(1) if m else None


def fmt_dt(dt, fmt_py="%Y-%m-%d %H:%M", fmt_net="yyyy-MM-dd HH:mm"):
    """Format a date that may be either a Python datetime OR a .NET
    System.DateTime (which is what q.QuerySql returns in IronPython and which
    has no Python strftime)."""
    if dt is None:
        return ""
    try:
        return dt.strftime(fmt_py)
    except:
        pass
    try:
        return dt.ToString(fmt_net)
    except:
        return str(dt)


def fmt_amount(cents):
    """transaction_amount is in cents (integer)."""
    try:
        return "${:,.2f}".format(float(cents) / 100.0)
    except:
        return ""


def build_record(row):
    """Turn one FortisTransactionLog row into a flat dict for display."""
    request_text = getattr(row, "Request", "") or ""
    friendly_type, url = extract_endpoint(request_text)
    body, raw_body = extract_body_json(request_text)
    detail = get_error_detail(row)
    cat = categorize_row(row)
    status = get_http_status(row)

    # Defaults
    name = ""
    last4 = ""
    street = ""
    phone = ""
    city = state = zipc = ""
    amount = ""
    exp = ""
    routing = ""
    people_id = None

    if body:
        addr = body.get("billing_address") or body.get("address") or {}
        street = addr.get("street", "") or ""
        phone = addr.get("phone", "") or ""
        city = addr.get("city", "") or ""
        state = addr.get("state", "") or ""
        zipc = addr.get("postal_code", "") or ""
        last4 = last4_from_account(body.get("account_number"))
        routing = body.get("routing_number", "") or ""
        exp = body.get("exp_date", "") or ""
        amount = fmt_amount(body.get("transaction_amount")) if body.get("transaction_amount") is not None else ""
        # Name: transactions use account_holder_name; contacts use first/last
        name = body.get("account_holder_name") or ""
        if not name:
            fn = body.get("first_name", "") or ""
            ln = body.get("last_name", "") or ""
            name = (fn + " " + ln).strip()
        people_id = people_id_from_contact_api(body)
    else:
        # Fallback: regex straight out of the raw body text
        m = re.search(r'"street"\s*:\s*"([^"]*)"', raw_body)
        if m:
            street = m.group(1)
        m = re.search(r'"phone"\s*:\s*"([^"]*)"', raw_body)
        if m:
            phone = m.group(1)
        m = re.search(r'"account_number"\s*:\s*"([^"]*)"', raw_body)
        if m:
            last4 = last4_from_account(m.group(1))
        m = re.search(r'"account_holder_name"\s*:\s*"([^"]*)"', raw_body)
        if m:
            name = m.group(1)

    # The "offending value" and its length depend on the category
    if cat == "street":
        value = street
    elif cat == "phone":
        value = phone
    elif cat == "routing":
        value = routing
    elif cat == "card":
        value = ("ending " + last4) if last4 else "(masked)"
    else:
        value = ""
    length = len(value) if (cat in ("street", "phone")) else ""

    return {
        "dt": getattr(row, "LogDateTime", None),
        "type": friendly_type,
        "category": cat,
        "status": status,
        "name": name,
        "last4": last4,
        "value": value,
        "length": length,
        "amount": amount,
        "exp": exp,
        "city": city,
        "state": state,
        "zip": zipc,
        "people_id": people_id,
        "detail": detail,
    }


#####################################################################
#### DATA ACCESS
#####################################################################

def sql_escape(s):
    return (s or "").replace("'", "''")


def get_rows(start_dt, end_dt, search_term):
    """Pull error rows in the date window (optionally text-filtered)."""
    where = "WHERE LogDateTime >= '{0}' AND LogDateTime < '{1}'".format(
        start_dt.strftime("%Y-%m-%d %H:%M:%S"),
        end_dt.strftime("%Y-%m-%d %H:%M:%S"),
    )
    if search_term:
        term = sql_escape(search_term)
        where += " AND (Request LIKE '%{0}%' OR FortisError LIKE '%{0}%' OR Response LIKE '%{0}%')".format(term)

    sql = """
        SELECT TOP {0}
            FortisTransactionLogId, LogDateTime, Request, Response,
            FortisError, ErrorMessage
        FROM dbo.FortisTransactionLog WITH (NOLOCK)
        {1}
        ORDER BY LogDateTime DESC
    """.format(MAX_DETAIL_ROWS, where)
    return q.QuerySql(sql)


#####################################################################
#### HEALTH DATA
#####################################################################
# Every timestamp in FortisWebhookTransactions is a Unix epoch stored as
# nvarchar, so each query guards with ISNUMERIC before casting.

EPOCH_UTC = "DATEADD(second, CAST(w.CreatedTS AS bigint), '1970-01-01')"
NUMERIC_TS = "ISNUMERIC(w.CreatedTS) = 1 AND w.CreatedTS NOT LIKE '%[^0-9]%'"


def to_pydt(dt):
    """Normalise a timestamp to a real Python datetime.

    q.QuerySql hands back .NET System.DateTime in IronPython. It has no
    strftime and will not add a Python timedelta, so anything that arrives
    from SQL has to be converted before it is formatted or shifted. Note that
    str() on one gives US format ("9/15/2026 1:34:08 PM"), not ISO, so the
    property path below is the reliable conversion rather than parsing text.
    """
    if dt is None:
        return None
    if isinstance(dt, datetime.datetime):
        return dt
    try:
        return datetime.datetime(dt.Year, dt.Month, dt.Day,
                                 dt.Hour, dt.Minute, dt.Second)
    except Exception:
        return None


def utc_to_local(dt):
    dt = to_pydt(dt)
    if dt is None:
        return None
    return dt + datetime.timedelta(hours=UTC_OFFSET_HOURS)


def fmt_local(dt, fallback="never"):
    """Format a UTC timestamp in local time, whatever type it arrived as."""
    local = utc_to_local(dt)
    if local is not None:
        return local.strftime("%Y-%m-%d %H:%M")
    if dt is None:
        return fallback
    # Unconvertible but not null: show what we have rather than lose it.
    return fmt_dt(dt)


def health_methods():
    """Payment methods seen in the last year, so the tab adapts if one is added."""
    sql = """
        SELECT DISTINCT LOWER(LTRIM(RTRIM(w.PaymentMethod))) AS Method
        FROM dbo.FortisWebhookTransactions w
        WHERE w.Event = 'CREATE' AND {0}
          AND {1} >= DATEADD(month, -12, GETUTCDATE())
          AND ISNULL(w.PaymentMethod, '') <> ''
    """.format(NUMERIC_TS, EPOCH_UTC)
    return [r.Method for r in q.QuerySql(sql)]


def health_heartbeat():
    """Last transaction and recent volume per payment method."""
    sql = """
        SELECT LOWER(LTRIM(RTRIM(w.PaymentMethod))) AS Method,
               MAX({0}) AS LastSeenUtc,
               COUNT(*) AS TotalEver,
               SUM(CASE WHEN {0} >= DATEADD(day, -30, GETUTCDATE()) THEN 1 ELSE 0 END) AS Last30,
               DATEDIFF(hour, MAX({0}), GETUTCDATE()) AS HoursAgo
        FROM dbo.FortisWebhookTransactions w
        WHERE w.Event = 'CREATE' AND {1} AND ISNULL(w.PaymentMethod, '') <> ''
        GROUP BY LOWER(LTRIM(RTRIM(w.PaymentMethod)))
    """.format(EPOCH_UTC, NUMERIC_TS)
    out = {}
    for r in q.QuerySql(sql):
        out[r.Method] = {
            "last_utc": r.LastSeenUtc,
            "hours_ago": int(r.HoursAgo or 0),
            "total": int(r.TotalEver or 0),
            "last30": int(r.Last30 or 0),
        }
    return out


def health_window_counts(methods):
    """Trailing window vs the same calendar window in prior months.

    Comparing like with like matters more than it sounds. Giving is strongly
    monthly: online volume triples on the 1st and ACH clusters there too, so a
    raw week-on-week comparison flags a normal mid-month lull as an outage.
    Anchoring to the same day-of-month range in previous months removes that.
    """
    today = datetime.datetime.utcnow()
    windows = []
    for k in range(0, HEALTH_BASELINE_MONTHS + 1):
        try:
            end = _months_back(today, k)
        except Exception:
            continue
        start = end - datetime.timedelta(days=HEALTH_WINDOW_DAYS - 1)
        windows.append({
            "label": "current" if k == 0 else start.strftime("%b"),
            "start": start.strftime("%Y-%m-%d 00:00:00"),
            "end": (end + datetime.timedelta(days=1)).strftime("%Y-%m-%d 00:00:00"),
            "is_current": k == 0,
            "range": "{0} to {1}".format(start.strftime("%b %d"), end.strftime("%b %d")),
        })

    for w in windows:
        sql = """
            SELECT LOWER(LTRIM(RTRIM(w.PaymentMethod))) AS Method,
                   COUNT(*) AS N,
                   ISNULL(SUM(CASE WHEN ISNUMERIC(w.TransactionAmount) = 1
                        THEN CAST(w.TransactionAmount AS money) ELSE 0 END), 0) AS Amt
            FROM dbo.FortisWebhookTransactions w
            WHERE w.Event = 'CREATE' AND {0}
              AND {1} >= '{2}' AND {1} < '{3}'
            GROUP BY LOWER(LTRIM(RTRIM(w.PaymentMethod)))
        """.format(NUMERIC_TS, EPOCH_UTC, w["start"], w["end"])
        counts = {}
        for r in q.QuerySql(sql):
            counts[r.Method] = {"n": int(r.N or 0), "amt": float(r.Amt or 0)}
        w["counts"] = counts
        for m in methods:
            w["counts"].setdefault(m, {"n": 0, "amt": 0.0})
    return windows


def _months_back(dt, k):
    """Same day-of-month k months earlier, clamped for short months."""
    month = dt.month - k
    year = dt.year
    while month <= 0:
        month += 12
        year -= 1
    day = dt.day
    while day > 1:
        try:
            return datetime.datetime(year, month, day)
        except ValueError:
            day -= 1
    return datetime.datetime(year, month, 1)


def health_assess(methods, heartbeat, windows):
    """Decide which methods look silent. Returns a list of per-method dicts."""
    current = None
    baseline = []
    for w in windows:
        if w["is_current"]:
            current = w
        else:
            baseline.append(w)

    rows = []
    for m in sorted(methods):
        hb = heartbeat.get(m, {})
        now_n = (current or {}).get("counts", {}).get(m, {}).get("n", 0)
        hist = [b["counts"].get(m, {}).get("n", 0) for b in baseline]
        expected = (sum(hist) / float(len(hist))) if hist else 0.0

        if now_n > 0:
            status, why = "ok", "{0} in the last {1} days".format(now_n, HEALTH_WINDOW_DAYS)
        elif expected >= HEALTH_MIN_EXPECTED:
            status = "silent"
            why = ("nothing in the last {0} days, but this window normally "
                   "produces about {1:.0f}".format(HEALTH_WINDOW_DAYS, expected))
        else:
            status = "quiet"
            why = ("nothing in the last {0} days, but this window is normally "
                   "near zero too ({1:.1f}), so this is not evidence of a "
                   "problem".format(HEALTH_WINDOW_DAYS, expected))

        rows.append({
            "method": m, "status": status, "why": why,
            "now": now_n, "expected": expected, "history": hist,
            "last_utc": hb.get("last_utc"), "hours_ago": hb.get("hours_ago", 0),
            "last30": hb.get("last30", 0),
        })
    return rows


def health_ach_pipeline():
    """ACH transactions that have not reached settled status."""
    sql = """
        WITH g AS (
            SELECT w.FortisWebHookTransactionId AS Tid,
                   MIN({0}) AS CreatedUtc,
                   MAX(CASE WHEN ISNUMERIC(w.StatusId) = 1
                            THEN CAST(w.StatusId AS int) ELSE 0 END) AS MaxStatus,
                   MAX(CASE WHEN ISNUMERIC(w.TransactionAmount) = 1
                            THEN CAST(w.TransactionAmount AS money) ELSE 0 END) AS Amt,
                   MAX(ISNULL(w.AccountHolderName, '')) AS Who
            FROM dbo.FortisWebhookTransactions w
            WHERE LOWER(w.PaymentMethod) = 'ach' AND {1}
              AND {0} >= DATEADD(day, -90, GETUTCDATE())
            GROUP BY w.FortisWebHookTransactionId
        )
        SELECT TOP 100 Tid, CreatedUtc, MaxStatus, Amt, Who,
               DATEDIFF(day, CreatedUtc, GETUTCDATE()) AS AgeDays
        FROM g
        WHERE MaxStatus < {2}
        ORDER BY CreatedUtc DESC
    """.format(EPOCH_UTC, NUMERIC_TS, SETTLED_STATUS)
    return list(q.QuerySql(sql))


def health_monthly(months=12):
    """Per-month volume, value and approval rate for each method."""
    sql = """
        SELECT CONVERT(varchar(7), {0}, 120) AS Mon,
               LOWER(LTRIM(RTRIM(w.PaymentMethod))) AS Method,
               COUNT(*) AS N,
               ISNULL(SUM(CASE WHEN ISNUMERIC(w.TransactionAmount) = 1
                    THEN CAST(w.TransactionAmount AS money) ELSE 0 END), 0) AS Amt,
               SUM(CASE WHEN w.ReasonCodeId = '1000' THEN 1 ELSE 0 END) AS Approved
        FROM dbo.FortisWebhookTransactions w
        WHERE w.Event = 'CREATE' AND {1}
          AND {0} >= DATEADD(month, -{2}, GETUTCDATE())
          AND ISNULL(w.PaymentMethod, '') <> ''
        GROUP BY CONVERT(varchar(7), {0}, 120), LOWER(LTRIM(RTRIM(w.PaymentMethod)))
        ORDER BY 1, 2
    """.format(EPOCH_UTC, NUMERIC_TS, int(months))
    return list(q.QuerySql(sql))


def health_monthly_errors(months=12):
    """Denials per month, split by endpoint, to pair with volume.

    A raw error count is close to meaningless on its own. The 412s here run at
    a steady few percent of card volume month after month; as a count they
    look alarming, as a rate they are flat.
    """
    sql = """
        SELECT CONVERT(varchar(7), l.LogDateTime, 120) AS Mon,
               CASE WHEN l.Request LIKE '%/ach/%' THEN 'ach'
                    WHEN l.Request LIKE '%/cc/%' THEN 'cc'
                    ELSE 'other' END AS Kind,
               COUNT(*) AS N
        FROM dbo.FortisTransactionLog l
        WHERE l.LogDateTime >= DATEADD(month, -{0}, GETDATE())
        GROUP BY CONVERT(varchar(7), l.LogDateTime, 120),
                 CASE WHEN l.Request LIKE '%/ach/%' THEN 'ach'
                      WHEN l.Request LIKE '%/cc/%' THEN 'cc'
                      ELSE 'other' END
        ORDER BY 1, 2
    """.format(int(months))
    return list(q.QuerySql(sql))


def denial_denominators(start_dt, end_dt):
    """Transactions per method in the same window, for denial rates.

    Cheap: FortisWebhookTransactions is about 32k rows and a full twelve month
    aggregate returns in tens of milliseconds, so this adds one trivial query
    rather than anything worth caching.
    """
    sql = """
        SELECT LOWER(LTRIM(RTRIM(w.PaymentMethod))) AS Method, COUNT(*) AS N
        FROM dbo.FortisWebhookTransactions w
        WHERE w.Event = 'CREATE' AND {0}
          AND {1} >= '{2}' AND {1} < '{3}'
          AND ISNULL(w.PaymentMethod, '') <> ''
        GROUP BY LOWER(LTRIM(RTRIM(w.PaymentMethod)))
    """.format(NUMERIC_TS, EPOCH_UTC,
               start_dt.strftime("%Y-%m-%d %H:%M:%S"),
               end_dt.strftime("%Y-%m-%d %H:%M:%S"))
    out = {}
    try:
        for r in q.QuerySql(sql):
            out[r.Method] = int(r.N or 0)
    except Exception:
        pass
    out["all"] = sum(out.values())
    return out


def denial_rate_trend(months=6):
    """Monthly denial rate, so the cards can carry a direction as well as a number."""
    try:
        vol = health_monthly(months)
        errs = health_monthly_errors(months)
    except Exception:
        return []

    by_mon = {}
    for r in vol:
        d = by_mon.setdefault(r.Mon, {"txn": 0, "err": 0, "by": {}})
        d["txn"] += int(r.N or 0)
        d["by"][r.Method] = int(r.N or 0)
    for e in errs:
        d = by_mon.setdefault(e.Mon, {"txn": 0, "err": 0, "by": {}})
        d["err"] += int(e.N or 0)
        d.setdefault("errby", {})[e.Kind] = int(e.N or 0)

    out = []
    for mon in sorted(by_mon.keys()):
        d = by_mon[mon]
        out.append({
            "month": mon,
            "txn": d["txn"],
            "err": d["err"],
            "rate": (100.0 * d["err"] / d["txn"]) if d["txn"] else None,
        })
    return out


def render_rate_trend(trend):
    """A compact bar strip of denial rate by month."""
    usable = [t for t in trend if t["rate"] is not None]
    if len(usable) < 2:
        return ""
    peak = max(t["rate"] for t in usable) or 1.0
    latest = usable[-1]["rate"]
    prior = usable[-2]["rate"]
    if prior and latest > prior * 1.5:
        verdict = '<b style="color:#c0392b">rising</b>'
    elif prior and latest < prior * 0.67:
        verdict = '<b style="color:#27ae60">falling</b>'
    else:
        verdict = '<b>steady</b>'

    # Each column is a full-height flex box justified to the bottom, so the
    # percentage label always sits directly on top of its own bar. The earlier
    # version let bars grow into the label row above them.
    BAR_MAX = 52
    out = ['<div style="margin:10px 0 18px">']
    out.append('<div style="font-size:12px;color:#555;margin-bottom:8px">'
               'Denial rate by month &mdash; denials as a share of transactions '
               'attempted. Currently {0} at {1:.1f}%.</div>'.format(verdict, latest))
    out.append('<div class="ft-scroll"><div style="display:flex;align-items:flex-end;'
               'gap:8px;height:{0}px;min-width:{1}px">'.format(
                   BAR_MAX + 34, len(usable) * 56))
    for t in usable:
        h = max(3, int(BAR_MAX * t["rate"] / peak))
        colour = "#c0392b" if t is usable[-1] else "#8fa8c8"
        out.append(
            '<div style="flex:0 0 48px;display:flex;flex-direction:column;'
            'justify-content:flex-end;height:100%;text-align:center" '
            'title="{0}: {1} denials of {2} transactions">'
            '<div style="font-size:10px;color:#555;line-height:13px">{3:.1f}%</div>'
            '<div style="height:{4}px;background:{5};border-radius:2px 2px 0 0"></div>'
            '<div style="font-size:10px;color:#777;line-height:14px;'
            'border-top:1px solid #ddd;padding-top:1px">{6}</div>'
            '</div>'.format(h_escape(t["month"]), t["err"], t["txn"], t["rate"],
                            h, colour, h_escape(t["month"][5:])))
    out.append('</div></div></div>')
    return "".join(out)


#####################################################################
#### HEALTH ALERTING
#####################################################################
# State is machine-generated and strictly ASCII (method names, ISO
# timestamps, integers), so plain json is safe here. Do not start storing
# Fortis error text in it without switching to an ASCII-escaping encoder:
# json.dumps breaks on non-ASCII at the IronPython interop boundary.

def health_load_state():
    try:
        raw = model.TextContent(HEALTH_STATE_KEY)
        parsed = json.loads(raw) if raw else {}
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        return {}


def health_save_state(state):
    try:
        model.WriteContentText(HEALTH_STATE_KEY, json.dumps(state), "")
    except Exception:
        pass


def _hours_since(stamp):
    if not stamp:
        return 99999.0
    try:
        then = datetime.datetime.strptime(stamp, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return 99999.0
    d = datetime.datetime.utcnow() - then
    return (d.days * 86400.0 + d.seconds) / 3600.0


def health_send_alert(subject, body):
    if not HEALTH_ALERT_PEOPLE_ID:
        return False
    try:
        model.Email(HEALTH_ALERT_PEOPLE_ID, HEALTH_ALERT_PEOPLE_ID,
                    HEALTH_ALERT_FROM_EMAIL, HEALTH_ALERT_FROM_NAME, subject, body)
        return True
    except Exception:
        return False


def health_check_and_alert(assessment):
    """Email about newly silent methods, and once when they come back.

    Only "silent" alerts. "quiet" means the baseline says near-zero is normal
    for this window, and alerting on that would train everyone to ignore it.
    """
    state = health_load_state()
    alerted = state.get("alerted") or {}
    sent = []

    for row in assessment:
        m = row["method"]
        if row["status"] == "silent":
            if _hours_since(alerted.get(m)) >= HEALTH_ALERT_COOLDOWN_HOURS:
                body = (
                    "<p><b>{0}</b> has produced no transactions in the last {1} days.</p>"
                    "<p>{2}.</p>"
                    "<p>Last seen: {3} ({4} hours ago).<br>"
                    "Same window in previous months: {5}</p>"
                    "<p>Worth knowing: FortisTransactionLog records errors only, so a "
                    "stoppage like this leaves no error to find. Nothing being attempted "
                    "looks identical to nothing going wrong.</p>"
                ).format(
                    m.upper(), HEALTH_WINDOW_DAYS, row["why"],
                    fmt_local(row["last_utc"]),
                    row["hours_ago"],
                    ", ".join(str(h) for h in row["history"]) or "no history")
                if health_send_alert(
                        "Fortis: no {0} transactions in {1} days".format(
                            m.upper(), HEALTH_WINDOW_DAYS), body):
                    alerted[m] = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
                    sent.append(m)
        elif row["status"] == "ok" and alerted.get(m):
            health_send_alert(
                "Fortis: {0} transactions have resumed".format(m.upper()),
                "<p><b>{0}</b> is flowing again: {1}.</p>".format(m.upper(), row["why"]))
            alerted.pop(m, None)
            sent.append(m + " (recovered)")

    state["alerted"] = alerted
    state["last_check"] = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    health_save_state(state)
    return sent


def health_run_scheduled():
    """Entry point for the ScheduledTasks block. Never raises.

    TouchPoint's scheduled task runner emails admins on any unhandled
    exception, on every run, so this swallows everything.
    """
    try:
        methods = health_methods()
        if not methods:
            return
        hb = health_heartbeat()
        windows = health_window_counts(methods)
        health_check_and_alert(health_assess(methods, hb, windows))
    except Exception:
        pass


#####################################################################
#### INPUT HANDLING
#####################################################################

def get_param(name, default=""):
    try:
        v = getattr(model.Data, name)
        if v is None:
            return default
        return str(v).strip()
    except:
        return default


def parse_date(s, fallback):
    if not s:
        return fallback
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.datetime.strptime(s, fmt)
        except:
            continue
    return fallback


def resolve_window():
    """Figure out the date window from either preset days or explicit dates."""
    today = datetime.datetime.now()
    end_default = today + datetime.timedelta(days=1)  # inclusive of today

    start_s = get_param("start")
    end_s = get_param("end")
    days_s = get_param("days")

    if start_s or end_s:
        end_dt = parse_date(end_s, end_default)
        # if an explicit end date was given, make it inclusive
        if end_s:
            end_dt = end_dt + datetime.timedelta(days=1)
        start_dt = parse_date(start_s, end_dt - datetime.timedelta(days=DEFAULT_DAYS))
        return start_dt, end_dt

    # preset days
    try:
        days = int(days_s) if days_s else DEFAULT_DAYS
    except:
        days = DEFAULT_DAYS
    if days <= 0:
        days = DEFAULT_DAYS
    return end_default - datetime.timedelta(days=days), end_default


#####################################################################
#### HTML RENDERING
#####################################################################

def h_escape(s):
    if s is None:
        return ""
    s = unicode(s) if not isinstance(s, (str, unicode)) else s
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def render_styles():
    return """
    <style>
      .ft-wrap { font-family: 'Segoe UI', Arial, sans-serif; color:#222; max-width:1250px; }
      .ft-wrap h2 { margin-bottom:4px; }
      .ft-note { background:#f3f8ff; border-left:4px solid #2a72d4; padding:12px 16px;
                 margin:14px 0; font-size:13px; line-height:1.5; border-radius:4px; }
      .ft-note b { color:#1a4d8f; }
      .ft-filters { background:#f7f7f7; border:1px solid #e0e0e0; border-radius:6px;
                    padding:12px 16px; margin:14px 0; }
      .ft-filters label { font-size:12px; font-weight:600; margin-right:4px; }
      .ft-filters input, .ft-filters select { padding:5px 8px; font-size:13px;
                    border:1px solid #ccc; border-radius:4px; margin-right:12px; }
      .ft-btn { background:#2a72d4; color:#fff; border:none; padding:7px 16px;
                border-radius:4px; font-size:13px; cursor:pointer; }
      .ft-btn.secondary { background:#666; }
      .ft-cards { display:flex; flex-wrap:wrap; gap:10px; margin:16px 0; }
      .ft-card { border:1px solid #e0e0e0; border-radius:6px; padding:10px 14px;
                 min-width:150px; background:#fff; }
      .ft-card .n { font-size:26px; font-weight:700; }
      .ft-card .l { font-size:12px; color:#666; }
      .ft-card.active { outline:2px solid #2a72d4; }
      table.ft-tbl { border-collapse:collapse; width:100%; font-size:12px; margin-top:8px; }
      table.ft-tbl th, table.ft-tbl td { border:1px solid #ddd; padding:5px 7px; vertical-align:top; }
      table.ft-tbl th { background:#f0f0f0; text-align:left; position:sticky; top:0; }
      table.ft-tbl tr:nth-child(even) td { background:#fafafa; }
      .badge { display:inline-block; padding:1px 7px; border-radius:10px; font-size:11px; color:#fff; }
      .b-card{background:#c0392b} .b-street{background:#e67e22} .b-phone{background:#8e44ad}
      .b-routing{background:#16a085} .b-lastname{background:#7f8c8d} .b-other{background:#95a5a6}
      .b-datavalidation{background:#34495e}
      table.ft-tbl th.num, table.ft-tbl td.num { text-align:right; white-space:nowrap; }
      table.ft-tbl th { white-space:nowrap; }
      table.ft-tbl td.wrap { white-space:normal; }
      /* Health tables are wider than the denial table; let them scroll rather
         than run off the edge of the page. */
      .ft-scroll { overflow-x:auto; margin:8px 0 4px; }
      .ft-scroll table.ft-tbl { min-width:640px; }
      /* Group separators for the twelve month trend, where several columns
         belong to each payment method. */
      table.ft-tbl td.grp, table.ft-tbl th.grp { border-left:2px solid #bbb; }
      .mono { font-family:Consolas,monospace; }
      .len-bad { color:#c0392b; font-weight:700; }
      .detailmsg { color:#555; font-size:11px; max-width:360px; }
      .muted { color:#999; }
    </style>
    """


def render_filters(start_dt, end_dt, category, search, days_used):
    presets = [(30, "30d"), (90, "90d"), (180, "180d"), (365, "1yr")]
    preset_html = ""
    for d, lbl in presets:
        preset_html += '<a class="ft-btn secondary" style="text-decoration:none;margin-right:6px;" href="?days={0}">{1}</a>'.format(d, lbl)

    cat_opts = '<option value="">All categories</option>'
    for key, lbl in CATEGORIES:
        sel = ' selected' if category == key else ''
        cat_opts += '<option value="{0}"{1}>{2}</option>'.format(key, sel, h_escape(lbl))

    return """
    <div class="ft-filters">
      <form method="get" action="">
        <label>From</label>
        <input type="date" name="start" value="{start}">
        <label>To</label>
        <input type="date" name="end" value="{end}">
        <label>Category</label>
        <select name="category">{cats}</select>
        <label>Search (name / last4 / phone / street)</label>
        <input type="text" name="search" value="{search}" placeholder="e.g. 857030523 or 0093 or Iris Drive" size="26">
        <button class="ft-btn" type="submit">Filter</button>
      </form>
      <div style="margin-top:8px;">Quick range: {presets}
        <span class="muted" style="margin-left:10px;">showing {start} &rarr; {endincl}</span>
      </div>
    </div>
    """.format(
        start=start_dt.strftime("%Y-%m-%d"),
        end=(end_dt - datetime.timedelta(days=1)).strftime("%Y-%m-%d"),
        endincl=(end_dt - datetime.timedelta(days=1)).strftime("%Y-%m-%d"),
        cats=cat_opts,
        search=h_escape(search),
        presets=preset_html,
    )


def _rate_note(count, denom):
    """Sub-label showing a denial count as a rate, when there is an honest base."""
    if not denom:
        return ""
    return ('<div class="r" style="font-size:11px;color:#777;margin-top:2px">'
            '{0:.1f}% of {1:,}</div>').format(100.0 * count / denom, denom)


def render_cards(counts, category, base_qs, denoms=None):
    denoms = denoms or {}
    html = '<div class="ft-cards">'
    total = sum(counts.values())
    active = '' if category else ' active'
    html += ('<a class="ft-card{0}" style="text-decoration:none;color:inherit;" href="?{1}">'
             '<div class="n">{2}</div><div class="l">All denials</div>{3}</a>').format(
                 active, base_qs, total, _rate_note(total, denoms.get("all")))
    for key, lbl in CATEGORIES:
        c = counts.get(key, 0)
        active = ' active' if category == key else ''
        method = CATEGORY_DENOMINATOR.get(key)
        html += ('<a class="ft-card{0}" style="text-decoration:none;color:inherit;" '
                 'href="?{1}&category={2}"><div class="n">{3}</div>'
                 '<div class="l">{4}</div>{5}</a>').format(
                     active, base_qs, key, c, h_escape(lbl),
                     _rate_note(c, denoms.get(method)) if method else "")
    html += '</div>'
    return html


def render_table(records):
    badge_cls = {"card": "b-card", "street": "b-street", "phone": "b-phone",
                 "routing": "b-routing", "lastname": "b-lastname",
                 "datavalidation": "b-datavalidation", "other": "b-other"}
    html = '<table class="ft-tbl" id="ftTable"><thead><tr>'
    cols = ["Date (UTC)", "Type", "Category", "HTTP", "Name / Person",
            "Last 4", "Offending Value", "Len", "Amount", "Exp", "Fortis Message"]
    for c in cols:
        html += "<th>{0}</th>".format(c)
    html += "</tr></thead><tbody>"

    for r in records:
        dt = r["dt"]
        dt_s = fmt_dt(dt)
        cat = r["category"]
        badge = '<span class="badge {0}">{1}</span>'.format(
            badge_cls.get(cat, "b-other"), h_escape(CATEGORY_LABELS.get(cat, cat)))

        # Person cell
        if r["people_id"]:
            person = '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(
                r["people_id"], h_escape(r["name"] or ("PeopleId " + r["people_id"])))
        else:
            person = h_escape(r["name"])

        last4 = r["last4"] or ('<span class="muted">n/a</span>' if cat == "street" else "")

        # length highlighting for street>32 / phone!=10
        length = r["length"]
        len_html = ""
        if length != "":
            bad = (cat == "street" and length > 32) or (cat == "phone" and length != 10)
            len_html = '<span class="{0}">{1}</span>'.format("len-bad" if bad else "", length)

        html += "<tr>"
        html += "<td class='mono'>{0}</td>".format(dt_s)
        html += "<td>{0}</td>".format(h_escape(r["type"]))
        html += "<td>{0}</td>".format(badge)
        html += "<td>{0}</td>".format(h_escape(r["status"]))
        html += "<td>{0}</td>".format(person)
        html += "<td class='mono'>{0}</td>".format(last4)
        html += "<td class='mono'>{0}</td>".format(h_escape(r["value"]))
        html += "<td>{0}</td>".format(len_html)
        html += "<td>{0}</td>".format(h_escape(r["amount"]))
        html += "<td class='mono'>{0}</td>".format(h_escape(r["exp"]))
        html += "<td class='detailmsg'>{0}</td>".format(h_escape(r["detail"]))
        html += "</tr>"

    html += "</tbody></table>"
    return html


def render_actions_js():
    # CSV export + print popup (popup avoids TouchPoint print-CSS conflicts)
    return """
    <div style="margin:14px 0;">
      <button class="ft-btn" onclick="ftExportCsv()">Export CSV</button>
      <button class="ft-btn secondary" onclick="ftPrint()">Print</button>
    </div>
    <script>
    function ftTableData() {
      var t = document.getElementById('ftTable');
      var rows = [], trs = t.querySelectorAll('tr');
      for (var i=0;i<trs.length;i++){
        var cells = trs[i].querySelectorAll('th,td'), row=[];
        for (var j=0;j<cells.length;j++){ row.push(cells[j].innerText.replace(/\\s+/g,' ').trim()); }
        rows.push(row);
      }
      return rows;
    }
    function ftExportCsv() {
      var rows = ftTableData(), csv = [];
      for (var i=0;i<rows.length;i++){
        var line = rows[i].map(function(c){ return '"' + c.replace(/"/g,'""') + '"'; }).join(',');
        csv.push(line);
      }
      var blob = new Blob([csv.join('\\n')], {type:'text/csv'});
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'fortis_denials.csv';
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
    }
    function ftPrint() {
      var t = document.getElementById('ftTable');
      if (!t) return;
      var css = '*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}';
      css += 'body{font-family:Segoe UI,Arial,sans-serif;padding:16px;color:#222}';
      css += 'table{border-collapse:collapse;width:100%;font-size:11px}';
      css += 'th,td{border:1px solid #999;padding:4px 6px;text-align:left}';
      css += 'th{background:#eee}';
      css += 'h3{margin:0 0 10px}';
      var pw = window.open('', '_blank');
      if (!pw){ alert('Popup blocked - please allow popups'); return; }
      pw.document.write('<!DOCTYPE html><html><head><title>Fortis Denials</title>');
      pw.document.write('<style>'+css+'</style></head><body>');
      pw.document.write('<h3>Fortis Denial Report</h3>');
      pw.document.write(t.outerHTML);
      pw.document.write('</body></html>');
      pw.document.close(); pw.focus();
      setTimeout(function(){ pw.print(); }, 300);
    }
    </script>
    """


#####################################################################
#### BLOCKED BEFORE FORTIS
#####################################################################
# TouchPoint runs its own checks before a payment is sent to Fortis. A payment
# stopped there never reaches the gateway, so neither FortisTransactionLog nor
# the webhook stream has a row for it, and TouchPoint does not log the block
# either. This tab shows what evidence does exist: the settings that decide
# each check, the IP scores TouchPoint saved, who was on those IPs, and the
# log line written when a returning giver skips the IP check.
#
# The order and wording below follow GivingPaymentModel.cs (bvcms source).

GATES_DEFAULT_DAYS = 14

# An IP used by at least this many different signed-in users in a week is a
# shared network (the church office, a school, a phone carrier). One high
# score there blocks everyone on it.
GATES_SHARED_IP_USERS = 10

# Setting id -> default TouchPoint uses when the row is missing.
GATE_SETTING_DEFAULTS = {
    "UseRecaptcha": "false",
    "ValidateAddressInGiving": "true",
    "BlockedEmailDomainsGiving": "tempmail.com,guerrillamail.com,throwaway.email,mailinator.com,yopmail.com",
    "FraudScoreMaximumGiving": "100",
    "FailedConsecutiveGiftBlockEnabled": "false",
    "FailedConsecutiveGiftLockout": "3",
    "FailedConsecutiveGiftTimeframe": "30",
    "FailedConsecutiveGiftBlockTime": "60",
    "FailedConsecutiveGiftIPIgnoreList": "",
}

BYPASS_ACTIVITY = "Fraud score check bypassed for authenticated user with contribution history"


def gate_settings():
    vals = dict(GATE_SETTING_DEFAULTS)
    ids = ",".join("'%s'" % k for k in GATE_SETTING_DEFAULTS)
    for r in q.QuerySql("SELECT Id, Setting FROM dbo.Setting WHERE Id IN (%s)" % ids):
        if r.Setting is not None and str(r.Setting).strip() != "":
            vals[str(r.Id)] = str(r.Setting).strip()
    return vals


def _bool(s):
    return str(s).strip().lower() == "true"


def _int(s, default):
    try:
        return int(str(s).strip())
    except Exception:
        return default


def fraud_limit(settings):
    return _int(settings["FraudScoreMaximumGiving"], 100)


def ip_blocked(score, limit):
    # Same rule as NetworkFraudScoreService.AllowAddressForGiving: a limit of
    # 100 turns the check off, otherwise anything above the limit is refused.
    return limit != 100 and score is not None and int(score) > limit


def _sql_in(values):
    return ",".join("'%s'" % sql_escape(v) for v in values) or "''"


def ip_people(ips, since):
    """Signed-in users seen on each IP since `since`, newest first."""
    out = {}
    if not ips:
        return out
    sql = """
        SELECT al.ClientIp, u.PeopleId, p.Name, MAX(al.ActivityDate) AS LastSeen
        FROM dbo.ActivityLog al WITH (NOLOCK)
        JOIN dbo.Users u ON u.UserId = al.UserId
        LEFT JOIN dbo.People p ON p.PeopleId = u.PeopleId
        WHERE al.ActivityDate >= '{0}' AND al.ClientIp IN ({1})
        GROUP BY al.ClientIp, u.PeopleId, p.Name
        ORDER BY al.ClientIp, MAX(al.ActivityDate) DESC
    """.format(since.strftime("%Y-%m-%d"), _sql_in(ips))
    for r in q.QuerySql(sql):
        out.setdefault(str(r.ClientIp), []).append(r)
    return out


def shared_ip_counts(ips):
    """Distinct signed-in users per IP over the last 7 days."""
    out = {}
    if not ips:
        return out
    sql = """
        SELECT ClientIp, COUNT(DISTINCT UserId) AS Users
        FROM dbo.ActivityLog WITH (NOLOCK)
        WHERE ActivityDate >= DATEADD(day, -7, GETDATE())
          AND UserId IS NOT NULL AND ClientIp IN ({0})
        GROUP BY ClientIp
    """.format(_sql_in(ips))
    for r in q.QuerySql(sql):
        out[str(r.ClientIp)] = int(r.Users or 0)
    return out


def ip_scores(ips):
    out = {}
    if not ips:
        return out
    sql = ("SELECT Evidence, Score, LastUpdated FROM dbo.FraudScores "
           "WHERE Evidence IN ({0})".format(_sql_in(ips)))
    for r in q.QuerySql(sql):
        out[str(r.Evidence)] = r
    return out


def gate_find_people(term):
    term = (term or "").strip()
    if not term:
        return []
    if term.isdigit():
        where = "p.PeopleId = %d" % int(term)
    elif "@" in term:
        t = sql_escape(term)
        where = "(p.EmailAddress = '{0}' OR p.EmailAddress2 = '{0}')".format(t)
    elif " " in term:
        t = sql_escape(term)
        where = "(p.Name LIKE '%{0}%' OR p.Name2 LIKE '%{0}%')".format(t)
    else:
        # Match the start of a name, or "Ivan" finds every Sullivan.
        t = sql_escape(term)
        where = ("(p.FirstName LIKE '{0}%' OR p.LastName LIKE '{0}%' "
                 "OR p.NickName LIKE '{0}%')".format(t))
    exact = sql_escape(term)
    sql = """
        SELECT TOP 25 p.PeopleId, p.Name, p.EmailAddress, p.Age, p.PrimaryCity,
            bill.TransactionDate AS LastBillDate, bill.OrganizationName AS LastBillOrg,
            ISNULL(NULLIF(bill.Amt, 0), ABS(bill.Amtdue)) AS LastBillAmt,
            pay.TransactionDate AS LastPayDate, pay.Amt AS LastPayAmt,
            (SELECT COUNT(*) FROM dbo.Contribution c WITH (NOLOCK)
             WHERE c.PeopleId = p.PeopleId AND c.ContributionStatusId = 0) AS RecordedGifts,
            (SELECT MAX(u.LastLoginDate) FROM dbo.Users u WHERE u.PeopleId = p.PeopleId) AS LastLogin,
            (SELECT COUNT(*) FROM dbo.Users u WHERE u.PeopleId = p.PeopleId) AS Accounts
        FROM dbo.People p
        OUTER APPLY (
            SELECT TOP 1 t.TransactionDate, t.Amt, t.Amtdue, o.OrganizationName
            FROM dbo.[Transaction] t WITH (NOLOCK)
            JOIN dbo.TransactionPeople tp ON tp.Id = t.OriginalId
            LEFT JOIN dbo.Organizations o ON o.OrganizationId = t.OrgId
            WHERE tp.PeopleId = p.PeopleId AND {2}
              AND t.TransactionDate >= DATEADD(day, -90, GETDATE())
            ORDER BY t.TransactionDate DESC
        ) bill
        OUTER APPLY (
            SELECT TOP 1 t.TransactionDate, t.Amt
            FROM dbo.[Transaction] t WITH (NOLOCK)
            JOIN dbo.TransactionPeople tp ON tp.Id = t.OriginalId
            WHERE tp.PeopleId = p.PeopleId AND t.Approved = 1 AND t.Amt > 0 AND NOT {2}
              AND t.TransactionDate >= DATEADD(day, -90, GETDATE())
            ORDER BY t.TransactionDate DESC
        ) pay
        WHERE {0} AND p.IsDeceased = 0
        ORDER BY CASE WHEN p.Name = '{1}' OR p.FirstName = '{1}' OR p.LastName = '{1}'
                      OR p.NickName = '{1}' THEN 0 ELSE 1 END,
                 CASE WHEN bill.TransactionDate IS NOT NULL THEN 0 ELSE 1 END,
                 bill.TransactionDate DESC, p.Name2
    """.format(where, exact, FEE_TX_SQL)
    return list(q.QuerySql(sql))


def person_ips(people_id, since):
    """IPs this person used while signed in. Guest payments leave nothing."""
    sql = """
        SELECT al.ClientIp, MAX(al.ActivityDate) AS LastSeen, COUNT(*) AS Hits
        FROM dbo.ActivityLog al WITH (NOLOCK)
        JOIN dbo.Users u ON u.UserId = al.UserId
        WHERE u.PeopleId = {0} AND al.ActivityDate >= '{1}' AND al.ClientIp IS NOT NULL
        GROUP BY al.ClientIp
        ORDER BY MAX(al.ActivityDate) DESC
    """.format(int(people_id), since.strftime("%Y-%m-%d"))
    return list(q.QuerySql(sql))


# Pay-balance links log the transaction's LoginPeopleId before the payer. On
# a bill a staff member created, that is the STAFF member, so their name shows
# up from every home that pays one of their bills. Anyone who has created
# transactions for at least this many other people in the last year is
# treated as billing staff and their rows are resolved to the real payer.
BILLING_STAFF_MIN_PAYERS = 5

# Visits under one billing-staff name more than this many minutes apart are
# different payers, so they are kept as separate rows.
STAFF_VISIT_GAP_MINUTES = 45

PAY_STEPS = ("OnlineReg PaymentForm", "OnlineReg PayDueStart")

# Steps TouchPoint logs only when a registration or pay link finished.
DONE_STEPS = ("OnlineReg PayDueConfirm", "OnlineReg Confirm", "OnlineReg ZeroDue",
              "OnlineReg JoinedOrg")


def to_pydt_safe(dt):
    try:
        v = to_pydt(dt)
        if v is not None:
            return v
    except Exception:
        pass
    for fmt, n in (("%Y-%m-%d %H:%M:%S", 19), ("%Y-%m-%d %H:%M", 16)):
        try:
            return datetime.datetime.strptime(str(dt)[:n], fmt)
        except Exception:
            continue
    return None


def ip_visits(ip, since):
    """Every registration and pay-balance step logged from one IP.

    These pages log each step with the IP and an involvement, even for
    guests. Guest visits to the online giving page are not logged at all, so
    blocked gifts cannot be found this way.
    """
    sql = """
        SELECT TOP 2000 al.ActivityDate, al.PeopleId, al.OrgId, al.Activity, al.UserId,
            p.Name, o.OrganizationName, u.PeopleId AS UserPeopleId, up.Name AS UserName,
            CASE WHEN EXISTS (SELECT 1 FROM dbo.Contribution c WITH (NOLOCK)
                              WHERE c.PeopleId = u.PeopleId AND c.ContributionStatusId = 0)
                 THEN 1 ELSE 0 END AS UserHasGift
        FROM dbo.ActivityLog al WITH (NOLOCK)
        LEFT JOIN dbo.People p ON p.PeopleId = al.PeopleId
        LEFT JOIN dbo.Organizations o ON o.OrganizationId = al.OrgId
        LEFT JOIN dbo.Users u ON u.UserId = al.UserId
        LEFT JOIN dbo.People up ON up.PeopleId = u.PeopleId
        WHERE al.ClientIp = '{0}' AND al.ActivityDate >= '{1}'
          AND al.Activity LIKE 'OnlineReg%' AND al.PeopleId IS NOT NULL
        ORDER BY al.ActivityDate
    """.format(sql_escape(ip), since.strftime("%Y-%m-%d"))
    return list(q.QuerySql(sql))


# Staff add a fee or adjustment to someone's account, then open that
# person's pay link to run the card. The fee row is marked by its
# TransactionId ("AdjustFee (id)" or "Initial Tran (id)"), carries the staff
# member as LoginPeopleId, and names the real payer, so it is how a visit
# logged under staff is traced back to the family.
FEE_TX_SQL = ("(t.TransactionId LIKE 'AdjustFee%' OR t.TransactionId LIKE 'Initial Tran%' "
              "OR ISNULL(t.Message, '') LIKE 'ADJ%' OR ISNULL(t.Message, '') LIKE 'FEE%')")

# How long before a staff pay-link visit the fee row can be and still count
# as the bill they opened.
FEE_TO_VISIT_MINUTES = 15


def billing_staff(pids):
    pids = sorted(set(int(p) for p in pids if p))
    if not pids:
        return set()
    sql = """
        SELECT t.LoginPeopleId
        FROM dbo.[Transaction] t WITH (NOLOCK)
        JOIN dbo.TransactionPeople tp ON tp.Id = t.OriginalId
        WHERE t.LoginPeopleId IN ({0}) AND tp.PeopleId <> t.LoginPeopleId AND {1}
          AND t.TransactionDate >= DATEADD(day, -365, GETDATE())
        GROUP BY t.LoginPeopleId
        HAVING COUNT(DISTINCT tp.PeopleId) >= {2}
    """.format(",".join(str(p) for p in pids), FEE_TX_SQL, BILLING_STAFF_MIN_PAYERS)
    return set(int(r.LoginPeopleId) for r in q.QuerySql(sql))


def org_transactions(org_ids, since):
    """Fees and real payments in these involvements. A real payment is
    gateway-approved money in; fee rows are bills, not payments."""
    org_ids = sorted(set(int(o) for o in org_ids if o))
    if not org_ids:
        return []
    sql = """
        SELECT t.TransactionDate, t.OrgId, t.Amt, t.Amtdue, t.LoginPeopleId, t.OriginalId,
            tp.PeopleId AS PayerId, p.Name AS PayerName,
            CASE WHEN {2} THEN 1 ELSE 0 END AS IsFee
        FROM dbo.[Transaction] t WITH (NOLOCK)
        JOIN dbo.TransactionPeople tp ON tp.Id = t.OriginalId
        LEFT JOIN dbo.People p ON p.PeopleId = tp.PeopleId
        WHERE t.OrgId IN ({0}) AND t.Approved = 1
          AND t.TransactionDate >= DATEADD(day, -1, '{1}')
        ORDER BY t.TransactionDate
    """.format(",".join(str(o) for o in org_ids), since.strftime("%Y-%m-%d"), FEE_TX_SQL)
    out = []
    for r in q.QuerySql(sql):
        fee = bool(int(r.IsFee or 0))
        amt = float(r.Amt or 0)
        out.append({"at": to_pydt_safe(r.TransactionDate), "raw_at": r.TransactionDate,
                    "org": int(r.OrgId or 0), "amt": amt, "due": float(r.Amtdue or 0),
                    "login": int(r.LoginPeopleId or 0), "payer": int(r.PayerId or 0),
                    "payer_name": r.PayerName or "", "orig": int(r.OriginalId or 0),
                    "fee": fee, "paid": (not fee) and amt > 0})
    return out


def open_bills(staff_id, org_id, at):
    """People in this involvement with a bill from this staff member and no
    payment since `at`. One of them is who was on the payment page."""
    sql = """
        SELECT TOP 12 tp.PeopleId, p.Name, MAX(t.TransactionDate) AS LastBill
        FROM dbo.[Transaction] t WITH (NOLOCK)
        JOIN dbo.TransactionPeople tp ON tp.Id = t.OriginalId
        JOIN dbo.People p ON p.PeopleId = tp.PeopleId
        WHERE t.OrgId = {1} AND t.LoginPeopleId = {0} AND tp.PeopleId <> {0}
          AND t.TransactionDate >= DATEADD(day, -180, '{2}')
          AND EXISTS (SELECT 1 FROM dbo.OrganizationMembers om
                      WHERE om.OrganizationId = {1} AND om.PeopleId = tp.PeopleId)
          AND NOT EXISTS (
              SELECT 1 FROM dbo.[Transaction] t2 WITH (NOLOCK)
              JOIN dbo.TransactionPeople tp2 ON tp2.Id = t2.OriginalId
              WHERE tp2.PeopleId = tp.PeopleId AND t2.OrgId = {1}
                AND t2.Approved = 1 AND t2.Amt > 0 AND t2.TransactionDate >= '{2}'
                AND NOT {3})
        GROUP BY tp.PeopleId, p.Name
        ORDER BY MAX(t.TransactionDate) DESC
    """.format(int(staff_id), int(org_id), at.strftime("%Y-%m-%d %H:%M:%S"),
               FEE_TX_SQL.replace("t.", "t2."))
    return list(q.QuerySql(sql))


def group_visits(visits, staff):
    """One row per person and involvement, except under billing staff, where
    each burst of visits is a separate payer and gets its own row."""
    groups = {}
    order = []
    for v in visits:
        pid = int(v.PeopleId)
        org = int(v.OrgId or 0)
        at = to_pydt_safe(v.ActivityDate)
        key = (pid, org)
        if pid in staff:
            prev = groups.get(("open",) + key)
            if prev is None or at is None or prev["last"] is None or \
                    at - prev["last"] > datetime.timedelta(minutes=STAFF_VISIT_GAP_MINUTES):
                key = key + (len(order),)
                groups[("open", pid, org)] = None
            else:
                key = prev["key"]
        g = groups.get(key)
        if g is None:
            g = {"key": key, "pid": pid, "org": org, "name": v.Name or "",
                 "org_name": v.OrganizationName or "", "first": at, "last": at,
                 "raw_first": v.ActivityDate, "raw_last": v.ActivityDate,
                 "pay": False, "signed_in": False, "staff": pid in staff,
                 "user_name": "", "user_pid": 0, "skips_check": False, "done": False}
            groups[key] = g
            order.append(key)
        g["last"] = at
        g["raw_last"] = v.ActivityDate
        if v.Activity in PAY_STEPS:
            g["pay"] = True
        if v.Activity in DONE_STEPS:
            g["done"] = True
        if v.UserId is not None:
            g["signed_in"] = True
            g["user_name"] = v.UserName or ""
            g["user_pid"] = int(v.UserPeopleId or 0)
            # Signing in only skips the IP check when the signed-in account
            # has a recorded gift. Staff impersonating a member sign in AS
            # that member, so the member's gift history is what counts.
            if int(v.UserHasGift or 0):
                g["skips_check"] = True
        if pid in staff:
            groups[("open", pid, org)] = g
    return [groups[k] for k in order]


def _as_label(g):
    if not g["signed_in"]:
        return "guest"
    who = h_escape(g["user_name"]) or "someone"
    if g["skips_check"]:
        return 'signed in as {0} <span class="muted">(has a gift, skips the IP check)</span>'.format(who)
    return ('signed in as {0} <span style="color:#c0392b">(no recorded gift, still '
            'checked)</span>'.format(who))


def find_fee(g, txs):
    """The bill a staff member created just before opening a pay link."""
    best = None
    for t in txs:
        if not t["fee"] or t["org"] != g["org"] or t["login"] != g["pid"] \
                or t["at"] is None or g["first"] is None:
            continue
        if g["first"] - datetime.timedelta(minutes=FEE_TO_VISIT_MINUTES) <= t["at"] \
                <= g["first"] + datetime.timedelta(minutes=1):
            best = t
    return best


def match_payment(g, txs, fee=None):
    for t in txs:
        if not t["paid"] or t["org"] != g["org"] or t["at"] is None or g["first"] is None:
            continue
        if t["at"] < g["first"] - datetime.timedelta(minutes=1):
            continue
        if fee is not None:
            if t["orig"] == fee["orig"] or t["payer"] == fee["payer"]:
                return t
        elif not g["staff"] and (t["payer"] == g["pid"] or t["login"] == g["pid"]):
            return t
    return None


def family_pay_visits(people_id, since):
    """Registration and pay-link steps that may belong to this family, from
    any IP: logged under a family member, or under the staff member who
    created one of the family's bills."""
    sql = """
        ;WITH fam AS (
            SELECT PeopleId FROM dbo.People
            WHERE FamilyId = (SELECT FamilyId FROM dbo.People WHERE PeopleId = {0})
        ), tx AS (
            SELECT DISTINCT t.OrgId, t.LoginPeopleId
            FROM dbo.[Transaction] t WITH (NOLOCK)
            JOIN dbo.TransactionPeople tp ON tp.Id = t.OriginalId
            WHERE tp.PeopleId IN (SELECT PeopleId FROM fam)
              AND t.TransactionDate >= DATEADD(day, -400, GETDATE())
        ), orgs AS (
            SELECT OrgId FROM tx
            UNION SELECT om.OrganizationId FROM dbo.OrganizationMembers om
            WHERE om.PeopleId IN (SELECT PeopleId FROM fam)
        )
        SELECT TOP 60 al.ActivityDate, al.ClientIp, al.PeopleId, p.Name, al.OrgId,
            o.OrganizationName, al.Activity, al.UserId,
            CASE WHEN al.PeopleId IN (SELECT PeopleId FROM fam) THEN 1 ELSE 0 END AS IsFamily
        FROM dbo.ActivityLog al WITH (NOLOCK)
        LEFT JOIN dbo.People p ON p.PeopleId = al.PeopleId
        LEFT JOIN dbo.Organizations o ON o.OrganizationId = al.OrgId
        WHERE al.ActivityDate >= '{1}' AND al.Activity LIKE 'OnlineReg%'
          AND al.OrgId IN (SELECT OrgId FROM orgs)
          AND (al.PeopleId IN (SELECT PeopleId FROM fam)
               OR (al.Activity IN ('OnlineReg PayDueStart', 'OnlineReg PaymentForm')
                   AND EXISTS (SELECT 1 FROM tx WHERE tx.OrgId = al.OrgId
                               AND tx.LoginPeopleId = al.PeopleId)))
        ORDER BY al.ActivityDate DESC
    """.format(int(people_id), since.strftime("%Y-%m-%d"))
    return list(q.QuerySql(sql))


def render_ip_drill(ip, since, days, settings):
    limit = fraud_limit(settings)
    sc = ip_scores([ip]).get(ip)
    users = shared_ip_counts([ip]).get(ip, 0)
    out = []
    a = out.append
    a('<h3 style="margin-top:22px">Activity from {0}</h3>'.format(h_escape(ip)))
    a('<p>Score: {0}{1}. Signed-in users on it this week: <b>{2}</b>{3}. '
      '<a href="?tab=gates&amp;gdays={4}">Back to all IPs</a></p>'.format(
          _score_badge(sc.Score if sc else None, limit),
          (", saved " + fmt_dt(sc.LastUpdated)) if sc else "",
          users, " (shared network)" if users >= GATES_SHARED_IP_USERS else "", days))
    if sc and ip_blocked(sc.Score, limit) and users >= GATES_SHARED_IP_USERS:
        a('<div class="ft-note" style="border-left-color:#c0392b"><b>Everyone paying as a guest '
          'from this network is being refused.</b> The score is reused for 30 days from the time '
          'above. Signed-in givers with a recorded gift still get through. Raising '
          '<code>FraudScoreMaximumGiving</code> above {0} (or to 100, which turns the check off) '
          'clears it immediately.</div>'.format(int(sc.Score)))

    # A block happens from the moment the score is saved. LastUpdated is
    # church time on the first save but UTC on a refresh, so allow for both.
    blocked_ip = sc is not None and ip_blocked(sc.Score, limit)
    scored_at = to_pydt_safe(sc.LastUpdated) if sc else None
    cutoff = scored_at - datetime.timedelta(hours=6) if scored_at else None

    visits = ip_visits(ip, since)
    a('<h4 style="margin-top:16px">Registration and pay-balance attempts</h4>')
    if not visits:
        a('<p class="muted">None from this IP in the last {0} days.</p>'.format(days))
    else:
        staff = billing_staff([v.PeopleId for v in visits])
        groups = group_visits(visits, staff)
        txs = org_transactions([g["org"] for g in groups], since)
        groups.sort(key=lambda g: g["last"] or datetime.datetime.min, reverse=True)
        counts = {"blocked": 0, "unpaid": 0, "paid": 0}
        body = []
        for g in groups:
            # Signed in as the very person the row is logged under means it
            # is their own visit, even if they also bill other people.
            if g["staff"] and g["signed_in"] and g["user_pid"] == g["pid"]:
                g["staff"] = False
            g["fee"] = find_fee(g, txs) if g["staff"] else None
            g["paid_tx"] = match_payment(g, txs, g["fee"])
        for g in groups:
            fee, paid = g["fee"], g["paid_tx"]
            hot = ''
            if paid is not None:
                status = '<span style="color:#27ae60">Paid ${0:,.2f} on {1}</span>'.format(
                    paid["amt"], fmt_dt(paid["raw_at"]))
                counts["paid"] += 1
            elif g["done"]:
                status = '<span style="color:#27ae60">Completed</span>'
                counts["paid"] += 1
            elif g["pay"] and blocked_ip and not g["skips_check"] and \
                    (cutoff is None or g["last"] is None or g["last"] >= cutoff):
                status = ('<span style="background:#c0392b;color:#fff;padding:2px 8px;'
                          'border-radius:3px">Likely blocked</span>')
                counts["blocked"] += 1
                hot = ' style="background:#fdecea"'
            elif g["pay"]:
                status = '<span style="color:#b9770e">Reached payment, not paid</span>'
                counts["unpaid"] += 1
            else:
                status = '<span class="muted">Did not reach payment</span>'

            staff_link = '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(
                g["pid"], h_escape(g["name"]))
            if not g["staff"]:
                person = staff_link
            elif fee is not None:
                person = ('<a href="/Person2/{0}" target="_blank"><b>{1}</b></a>'
                          '<div style="font-size:11px;color:#777">{2} billed them ${3:,.2f} at '
                          '{4}, then opened their pay link.</div>'.format(
                              fee["payer"], h_escape(fee["payer_name"]), staff_link,
                              fee["amt"] or abs(fee["due"]), fmt_dt(fee["raw_at"])))
            else:
                person = ('<span class="muted">pay link for a bill from</span> ' + staff_link +
                          '<div style="font-size:11px;color:#777">Logged under the staff member '
                          'who created the bill; the payer is not recorded.</div>')
                if paid is None and g["pay"]:
                    # Staff often reopen a bill from an earlier day instead of
                    # making a new one, so there is no fresh bill to trace.
                    # Someone whose earlier try on a bill from the same staff
                    # and involvement is still unpaid is the likeliest payer.
                    probable = []
                    for e in groups:
                        if e is g or e["fee"] is None or e["paid_tx"] is not None:
                            continue
                        if e["pid"] == g["pid"] and e["org"] == g["org"] and \
                                e["first"] is not None and g["first"] is not None and \
                                e["first"] < g["first"] and \
                                e["fee"]["payer"] not in [x["fee"]["payer"] for x in probable]:
                            probable.append(e)
                    if probable:
                        person += ('<div style="font-size:12px;margin-top:3px"><b>Probably ' +
                                   ", ".join('<a href="/Person2/{0}" target="_blank">{1}</a> '
                                             '(billed ${2:,.2f} on {3}, still unpaid)'.format(
                                                 e["fee"]["payer"], h_escape(e["fee"]["payer_name"]),
                                                 e["fee"]["amt"] or abs(e["fee"]["due"]),
                                                 fmt_dt(e["fee"]["raw_at"])) for e in probable) +
                                   '</b>: an earlier try on a bill from the same person was '
                                   'never paid, and reopening it leaves no new bill to trace.</div>')
                    try:
                        cands = open_bills(g["pid"], g["org"], g["first"])
                    except Exception:
                        cands = []
                    cands = [c for c in cands if int(c.PeopleId) not in
                             [e["fee"]["payer"] for e in probable]]
                    if cands:
                        person += ('<div style="font-size:11px">Other open bills from them here, newest first: ' +
                                   ", ".join('<a href="/Person2/{0}" target="_blank">{1}</a>'.format(
                                       int(c.PeopleId), h_escape(c.Name or "")) for c in cands) +
                                   '</div>')
            body.append('<tr{0}><td class="wrap">{1}</td>'
                        '<td><a href="/Org/{2}" target="_blank">{3}</a></td>'
                        '<td>{4}</td><td>{5}</td><td>{6}</td><td class="wrap">{7}</td></tr>'.format(
                            hot, person, g["org"], h_escape(g["org_name"]),
                            fmt_dt(g["raw_first"]), fmt_dt(g["raw_last"]),
                            _as_label(g), status))
        a('<p><b>{0}</b> likely blocked, {1} reached payment but did not pay, {2} paid.</p>'.format(
            counts["blocked"], counts["unpaid"], counts["paid"]))
        a('<div class="ft-scroll"><table class="ft-tbl"><thead><tr><th>Person</th>'
          '<th>Involvement</th><th>First seen</th><th>Last seen</th><th>As</th><th>Result</th>'
          '</tr></thead><tbody>' + "".join(body) + '</tbody></table></div>')
    a('<p style="font-size:11px;color:#777">&ldquo;Likely blocked&rdquo; means a guest reached '
      'the payment form on this IP after it was scored, was not signed in as someone with a '
      'recorded gift, and no approved payment followed. Staff impersonating a member are '
      'signed in as that member, so the member&rsquo;s gifts decide it. A '
      'person who simply gave up looks the same. Pay-balance links are logged under whoever '
      'created the bill, which is often staff; those rows list the people with open bills from '
      'that staff member, and one of them was the payer. To check one family directly, use the '
      'person lookup. Guest visits to the online giving page are not logged by TouchPoint, so '
      'blocked gifts do not appear here.</p>')

    who = ip_people([ip], since).get(ip, [])
    a('<details style="margin-top:10px"><summary>Everyone signed in on this IP ({0})</summary>'.format(len(who)))
    if who:
        a('<table class="ft-tbl" style="width:auto"><thead><tr><th>Who</th><th>Last seen</th>'
          '</tr></thead><tbody>')
        for r in who[:200]:
            a('<tr><td><a href="/Person2/{0}" target="_blank">{1}</a></td><td>{2}</td></tr>'.format(
                int(r.PeopleId), h_escape(r.Name or ""), fmt_dt(r.LastSeen)))
        a('</tbody></table>')
    a('</details>')
    return "".join(out)


def _ip_link(ip, days):
    return '<a href="?tab=gates&amp;gdays={0}&amp;ip={1}">{2}</a>'.format(
        days, h_escape(ip), h_escape(ip))


def _score_badge(score, limit):
    if score is None:
        return '<span class="muted">not scored</span>'
    if ip_blocked(score, limit):
        return ('<span style="background:#c0392b;color:#fff;padding:2px 8px;'
                'border-radius:3px">{0} BLOCKED</span>'.format(int(score)))
    return '<span style="color:#27ae60"><b>{0}</b> allowed</span>'.format(int(score))


def _names(rows, limit=3):
    if not rows:
        return '<span class="muted">nobody signed in</span>'
    shown = ", ".join('<a href="/Person2/{0}" target="_blank">{1}</a>'.format(
        int(r.PeopleId), h_escape(r.Name or "")) for r in rows[:limit])
    if len(rows) > limit:
        shown += ' <span class="muted">(+{0} more)</span>'.format(len(rows) - limit)
    return shown


def render_gate_table(s):
    limit = fraud_limit(s)
    lock_on = _bool(s["FailedConsecutiveGiftBlockEnabled"])
    rows = [
        ("Captcha", "Captcha must be completed",
         "On" if _bool(s["UseRecaptcha"]) else "Off",
         "Skipped for a signed-in giver with a recorded gift. Wallet payments skip it."),
        ("Billing address", "Invalid billing address",
         "On" if _bool(s["ValidateAddressInGiving"]) else "Off",
         "Only for card and bank entry, not wallets."),
        ("Email domain", "The email address provided cannot be used for online giving.",
         h_escape(s["BlockedEmailDomainsGiving"]) or "none",
         "Checks the billing email against this list."),
        ("IP fraud score", "Giving is not supported from your current location.",
         ("<b>Off</b> (limit 100)" if limit == 100 else
          '<b style="color:#c0392b">On</b>: scores above <b>{0}</b> are refused'.format(limit)),
         "Each IP gets a 0 to 100 risk score. Skipped for a signed-in giver with a "
         "recorded gift. Saved scores are reused for 30 days."),
        ("Failed-attempt lockout", "Giving is temporarily blocked due to too many failed attempts.",
         ("On: {0} failures in {1} min locks for {2} min".format(
             s["FailedConsecutiveGiftLockout"], s["FailedConsecutiveGiftTimeframe"],
             s["FailedConsecutiveGiftBlockTime"]) if lock_on else "Off"),
         "Counts are kept in TouchPoint's cache, not the database, so this page "
         "cannot list who is locked out. Ignored IPs: {0}".format(
             h_escape(s["FailedConsecutiveGiftIPIgnoreList"]) or "none")),
    ]
    out = ['<div class="ft-scroll"><table class="ft-tbl"><thead><tr>'
           '<th>Check</th><th>What the giver sees</th><th>Your setting</th>'
           '<th style="width:38%">Notes</th></tr></thead><tbody>']
    for name, msg, setting, note in rows:
        out.append('<tr><td><b>{0}</b></td><td class="wrap">&ldquo;{1}&rdquo;</td>'
                   '<td class="wrap">{2}</td><td class="wrap" style="font-size:12px">{3}'
                   '</td></tr>'.format(name, h_escape(msg), setting, note))
    out.append('</tbody></table></div>')
    return "".join(out)


def render_person_lookup(term, since, settings, days=GATES_DEFAULT_DAYS):
    limit = fraud_limit(settings)
    blocked_domains = [d.strip().lower() for d in settings["BlockedEmailDomainsGiving"].split(",") if d.strip()]
    people = gate_find_people(term)
    if not people:
        return '<p>No one found for &ldquo;{0}&rdquo;.</p>'.format(h_escape(term))
    if len(people) > 1:
        # Full detail for every match is a wall of tables; pick one first.
        rows = []
        for p in people:
            billed = to_pydt_safe(p.LastBillDate) if p.LastBillDate is not None else None
            paid = to_pydt_safe(p.LastPayDate) if p.LastPayDate is not None else None
            if billed is None:
                bill_txt, hot = '<span class="muted">none in 90 days</span>', ''
            else:
                unpaid = paid is None or paid < billed
                bill_txt = '{0}: ${1:,.2f} on {2}{3}'.format(
                    h_escape(p.LastBillOrg or ""), float(p.LastBillAmt or 0),
                    fmt_dt(p.LastBillDate),
                    ' <b style="color:#c0392b">not paid since</b>' if unpaid else '')
                hot = ' style="background:#fdecea"' if unpaid else ''
            pay_txt = ('${0:,.2f} on {1}'.format(float(p.LastPayAmt or 0), fmt_dt(p.LastPayDate))
                       if p.LastPayDate is not None else '<span class="muted">none in 90 days</span>')
            rows.append(
                '<tr{0}><td><a href="?tab=gates&amp;gdays={1}&amp;person={2}"><b>{3}</b></a></td>'
                '<td class="num">{2}</td><td class="num">{4}</td><td>{5}</td><td>{6}</td>'
                '<td class="wrap">{7}</td><td>{8}</td></tr>'.format(
                    hot, days, int(p.PeopleId), h_escape(p.Name),
                    int(p.Age) if p.Age is not None else "", h_escape(p.PrimaryCity or ""),
                    h_escape(p.EmailAddress or ""), bill_txt, pay_txt))
        return ('<p>{0} people match &ldquo;{1}&rdquo;. Pick one. Anyone billed recently '
                'is listed first, and a red row has a bill with no payment since.</p>'
                '<div class="ft-scroll"><table class="ft-tbl"><thead><tr><th>Name</th>'
                '<th class="num">PeopleId</th><th class="num">Age</th><th>City</th><th>Email</th>'
                '<th>Last bill (90 days)</th><th>Last payment</th></tr></thead>'
                '<tbody>{2}</tbody></table></div>'.format(len(people), h_escape(term), "".join(rows)))

    out = []
    for p in people:
        gifts = int(p.RecordedGifts or 0)
        email = p.EmailAddress or ""
        domain = email.split("@")[-1].strip().lower() if "@" in email else ""
        out.append('<div style="border:1px solid #e0e0e0;border-radius:6px;padding:10px 14px;margin:10px 0">')
        out.append('<b><a href="/Person2/{0}" target="_blank">{1}</a></b> '
                   '<span class="muted">({0}) {2}</span>'.format(
                       int(p.PeopleId), h_escape(p.Name), h_escape(email)))

        facts = []
        if int(p.Accounts or 0) == 0:
            facts.append("No user account, so they always pay as a guest and get every check.")
        elif gifts > 0:
            facts.append("Has {0} recorded gift(s). <b>If they sign in first</b>, the captcha and "
                         "IP checks are skipped.".format(gifts))
        else:
            facts.append("Has an account but no recorded gifts, so signing in does "
                         "<b>not</b> skip the IP check.")
        if int(p.Accounts or 0):
            facts.append("Last sign in: {0}".format(fmt_dt(p.LastLogin) if p.LastLogin else "never"))
            try:
                links = account_signin_links(p.PeopleId, since)
            except Exception:
                links = []
            if links:
                facts.append("Sign-in links used for their account (staff impersonating them, or "
                             "an emailed sign-in link): " + ", ".join(
                                 fmt_dt(l.Expires) for l in links) +
                             ". While someone is signed in as them, the IP check uses <b>their</b> "
                             "gift history.")
        if domain and domain in blocked_domains:
            facts.append('<b style="color:#c0392b">Their email domain ({0}) is on the blocked '
                         'list.</b>'.format(h_escape(domain)))
        out.append('<ul style="margin:6px 0 6px 18px;font-size:13px">' +
                   "".join("<li>%s</li>" % f for f in facts) + '</ul>')

        ips = person_ips(p.PeopleId, since)
        if not ips:
            out.append('<p class="muted" style="font-size:12px">No signed-in activity in this '
                       'window, so there is no IP to check. A guest payment records nothing '
                       'that ties an IP to a person.</p>')
        else:
            ip_list = [str(r.ClientIp) for r in ips]
            scores = ip_scores(ip_list)
            shared = shared_ip_counts(ip_list)
            out.append('<table class="ft-tbl" style="width:auto"><thead><tr><th>IP they used</th>'
                       '<th>Last seen</th><th class="num">Users this week</th><th>Score</th>'
                       '</tr></thead><tbody>')
            for r in ips[:10]:
                ip = str(r.ClientIp)
                sc = scores.get(ip)
                users = shared.get(ip, 0)
                tag = (' <span class="muted">shared network</span>'
                       if users >= GATES_SHARED_IP_USERS else '')
                out.append('<tr><td>{0}{1}</td><td>{2}</td><td class="num">{3}</td>'
                           '<td>{4}</td></tr>'.format(
                               _ip_link(ip, days), tag, fmt_dt(r.LastSeen), users,
                               _score_badge(sc.Score if sc else None, limit)))
            out.append('</tbody></table>')
        out.append(render_family_visits(p.PeopleId, since, days, limit))
        out.append('</div>')
    return "".join(out)


def account_signin_links(people_id, since):
    """One-time sign-in links used for this person's accounts.

    Impersonating someone creates one of these (expiring after a minute) and
    then signs in as them, so a used link on the day of a failed payment is
    good evidence staff were in their account. Emailed sign-in links use the
    same table and look the same; who used the link is not stored.
    """
    sql = """
        SELECT TOP 20 u.Username, l.expires AS Expires
        FROM dbo.OneTimeLinks l WITH (NOLOCK)
        JOIN dbo.Users u ON u.Username = l.querystring
        WHERE u.PeopleId = {0} AND l.used = 1 AND l.expires >= '{1}'
        ORDER BY l.expires DESC
    """.format(int(people_id), since.strftime("%Y-%m-%d"))
    return list(q.QuerySql(sql))


def family_ids(people_id):
    sql = """
        SELECT PeopleId FROM dbo.People
        WHERE FamilyId = (SELECT FamilyId FROM dbo.People WHERE PeopleId = {0})
    """.format(int(people_id))
    return set(int(r.PeopleId) for r in q.QuerySql(sql))


def render_family_visits(people_id, since, days, limit):
    visits = family_pay_visits(people_id, since)
    if not visits:
        return ('<p class="muted" style="font-size:12px">No registration or pay-link visits '
                'for this family in the window.</p>')
    fam = family_ids(people_id)
    ips = sorted(set(str(v.ClientIp) for v in visits if v.ClientIp))
    scores = ip_scores(ips)
    staff_orgs = [v.OrgId for v in visits if not int(v.IsFamily or 0)]
    # Bills can predate the window by months, so look further back for them.
    txs = org_transactions(staff_orgs, since - datetime.timedelta(days=180)) if staff_orgs else []

    def family_owed_staff(g):
        """True when this family had an unpaid bill from this staff member in
        this involvement at the time of the visit."""
        last_bill = None
        for t in txs:
            if t["org"] == g["org"] and t["payer"] in fam and t["at"] is not None \
                    and g["first"] is not None and t["at"] <= g["first"]:
                if t["fee"] and t["login"] == g["pid"]:
                    last_bill = t["at"]
                elif t["paid"] and last_bill is not None and t["at"] >= last_bill:
                    last_bill = None
        return last_bill is not None

    # A visit logged under staff is a pay link for one of the bills they
    # created. The bill they created just before it names the family, so it
    # is either provably theirs or provably someone else's. Only visits with
    # no such bill stay uncertain, and of those only blocked IPs are kept.
    rows, other_family, uncertain_hidden = [], 0, 0
    for v in visits:
        sc = scores.get(str(v.ClientIp or ""))
        if int(v.IsFamily or 0):
            rows.append((v, sc, None, None))
            continue
        g = {"org": int(v.OrgId or 0), "pid": int(v.PeopleId), "staff": True,
             "first": to_pydt_safe(v.ActivityDate)}
        fee = find_fee(g, txs)
        if fee is not None:
            if fee["payer"] in fam:
                rows.append((v, sc, fee, match_payment(g, txs, fee)))
            else:
                other_family += 1
        elif sc is not None and ip_blocked(sc.Score, limit) and family_owed_staff(g):
            rows.append((v, sc, None, None))
        else:
            uncertain_hidden += 1
    if not rows:
        return ('<p class="muted" style="font-size:12px">No registration or pay-link visits for '
                'this family in the window.</p>')

    out = ['<p style="font-size:13px;margin:10px 0 4px"><b>Registration and pay-link visits '
           'for this family</b> (any IP, signed in or not)</p>'
           '<table class="ft-tbl"><thead><tr><th>When</th><th>IP</th>'
           '<th>Score</th><th>Involvement</th><th>Step</th><th>Who</th>'
           '</tr></thead><tbody>']
    for v, sc, fee, paid in rows:
        ip = str(v.ClientIp or "")
        blocked = sc is not None and ip_blocked(sc.Score, limit)
        staff_link = '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(
            int(v.PeopleId), h_escape(v.Name or ""))
        if int(v.IsFamily or 0):
            who = staff_link
        elif fee is not None:
            if paid is not None:
                result = '<span style="color:#27ae60">paid ${0:,.2f} at {1}</span>'.format(
                    paid["amt"], fmt_dt(paid["raw_at"]))
            elif blocked:
                result = ('<b style="color:#c0392b">no payment followed; this IP is blocked, '
                          'so this is very likely the refused payment</b>')
            else:
                result = 'no payment followed'
            who = ('<b>Theirs.</b> {0} billed <a href="/Person2/{1}" target="_blank">{2}</a> '
                   '${3:,.2f} at {4}, then opened the pay link; {5}.'.format(
                       staff_link, fee["payer"], h_escape(fee["payer_name"]),
                       fee["amt"] or abs(fee["due"]), fmt_dt(fee["raw_at"]), result))
        else:
            who = staff_link + (' <span class="muted">(pay link for a bill they created. This '
                                'family owed them one at the time, but so may others.)</span>')
        hot = ' style="background:#fdecea"' if blocked and v.Activity in PAY_STEPS \
            and (fee is None or paid is None) else ''
        out.append('<tr{0}><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td>'
                   '<td class="wrap">{6}</td></tr>'.format(
                       hot, fmt_dt(v.ActivityDate), _ip_link(ip, days) if ip else "",
                       _score_badge(sc.Score if sc else None, limit),
                       h_escape(v.OrganizationName or ""),
                       h_escape((v.Activity or "").replace("OnlineReg ", "")), who))
    out.append('</tbody></table>')
    if other_family or uncertain_hidden:
        out.append('<p class="muted" style="font-size:11px">Left out: {0} pay-link visits that '
                   'were other families&rsquo; bills from the same staff, and {1} that could not be '
                   'traced to this family.</p>'.format(other_family, uncertain_hidden))
    return "".join(out)


def internal_networks():
    """IPs that many different people signed in from this week: the church
    building, most likely. Returned busiest first."""
    sql = """
        SELECT TOP 3 ClientIp, COUNT(DISTINCT UserId) AS Users
        FROM dbo.ActivityLog WITH (NOLOCK)
        WHERE ActivityDate >= DATEADD(day, -7, GETDATE())
          AND UserId IS NOT NULL AND ClientIp IS NOT NULL
        GROUP BY ClientIp
        HAVING COUNT(DISTINCT UserId) >= {0}
        ORDER BY COUNT(DISTINCT UserId) DESC
    """.format(GATES_SHARED_IP_USERS)
    return [(str(r.ClientIp), int(r.Users or 0)) for r in q.QuerySql(sql)]


def ignored_ips(settings):
    raw = settings.get("FailedConsecutiveGiftIPIgnoreList") or ""
    return [x for x in re.split(r"[\s,;]+", raw) if x]


def render_findings(settings, days):
    """Plain-language list of what is wrong, worst first."""
    limit = fraud_limit(settings)
    lock_on = _bool(settings["FailedConsecutiveGiftBlockEnabled"])
    ignored = ignored_ips(settings)
    nets = internal_networks()
    scores = ip_scores([ip for ip, _ in nets])
    found = []  # (level, html)

    for ip, users in nets:
        sc = scores.get(ip)
        label = '{0} ({1} people signed in from it this week)'.format(_ip_link(ip, days), users)
        if sc is not None and ip_blocked(sc.Score, limit):
            scored_at = to_pydt_safe(sc.LastUpdated)
            until = (scored_at + datetime.timedelta(days=30)).strftime("%b %d") if scored_at else "30 days after it was scored"
            found.append(("bad",
                '<b>Your building&rsquo;s network is blocked.</b> {0} looks like an internal IP, '
                'and it scored <b>{1}</b> against your limit of {2} on {3}. Every guest payment '
                'made from the building (front desk, kiosks, staff helping someone pay) is '
                'refused with &ldquo;Giving is not supported from your current location&rdquo; '
                'until about <b>{4}</b>, and it may score high again after that.'
                '<ul style="margin:6px 0 0 18px">'
                '<li>Quickest: have staff <b>sign in</b> before running a payment. Anyone '
                'signed in who has a recorded gift of their own skips this check.</li>'
                '<li>Set <code>FraudScoreMaximumGiving</code> to 100 to turn the check off, or '
                'ask TouchPoint support to clear the saved score for this IP.</li>'
                '<li>Ask IT why the building scores high (a VPN, proxy, guest wifi, or an '
                'infected device on the network can all do it).</li>'
                '<li>The Failed Transaction Ignored IPs setting does <b>not</b> help here; it '
                'only covers the lockout below.</li></ul>'.format(
                    label, int(sc.Score), limit, fmt_dt(sc.LastUpdated), until)))
        elif sc is not None and limit != 100 and int(sc.Score) >= limit - 10:
            found.append(("warn",
                '{0} looks like an internal IP and scored <b>{1}</b>, close to your limit of '
                '{2}. A small change in its reputation will block every guest payment made '
                'from the building.'.format(label, int(sc.Score), limit)))

        if lock_on and ip not in ignored:
            found.append(("warn",
                '{0} looks like an internal IP but is not in <b>Failed Transaction Ignored '
                'IPs</b> (Admin &gt; Advanced &gt; Settings &gt; Finance &gt; Online Giving). '
                'A few declined cards at the front desk would lock everyone in the building out '
                'of paying for {1} minutes. Adding it does not affect the location check.'.format(
                    label, settings["FailedConsecutiveGiftBlockTime"])))
        elif lock_on:
            found.append(("ok",
                '{0} is in Failed Transaction Ignored IPs, so the building cannot be locked out '
                'by repeated declines. That list does not cover the location check.'.format(label)))

    stale = [x for x in ignored if x not in [ip for ip, _ in nets]]
    if stale:
        found.append(("info",
            'Failed Transaction Ignored IPs also lists {0}, which nobody signed in from this '
            'week. If the building&rsquo;s IP changed, update the list.'.format(
                ", ".join(h_escape(x) for x in stale))))

    if not nets:
        found.append(("info",
            'No internal network detected (no IP had {0}+ people signed in this week).'.format(
                GATES_SHARED_IP_USERS)))

    if not found:
        return ''
    colors = {"bad": "#c0392b", "warn": "#f0ad4e", "info": "#2a72d4", "ok": "#27ae60"}
    rank = {"bad": 0, "warn": 1, "info": 2, "ok": 3}
    found.sort(key=lambda f: rank[f[0]])
    out = ['<h3 style="margin-top:18px">Detected problems</h3>']
    for level, html in found:
        out.append('<div class="ft-note" style="border-left-color:{0};font-size:13px">{1}'
                   '</div>'.format(colors[level], html))
    return "".join(out)


def render_gates():
    try:
        days = int(get_param("gdays") or GATES_DEFAULT_DAYS)
    except Exception:
        days = GATES_DEFAULT_DAYS
    if days <= 0:
        days = GATES_DEFAULT_DAYS
    since = datetime.datetime.now() - datetime.timedelta(days=days)
    person = get_param("person")
    drill_ip = get_param("ip")
    if not re.match(r"^[0-9a-fA-F\.:]{3,45}$", drill_ip or ""):
        drill_ip = ""
    settings = gate_settings()
    limit = fraud_limit(settings)

    if drill_ip:
        return render_ip_drill(drill_ip, since, days, settings)

    out = []
    a = out.append
    a('<div class="ft-note"><b>Why the other tabs show nothing.</b> TouchPoint runs the '
      'checks below <b>before</b> a payment is sent to Fortis. A payment stopped here never '
      'reaches Fortis, so it is not in the Denials or Health tabs, and TouchPoint does not '
      'log the block itself. This tab collects the evidence that does exist. These checks '
      'apply to gifts and to registration payments.</div>')

    a('<div class="ft-filters"><form method="get" action="">'
      '<input type="hidden" name="tab" value="gates">'
      '<label>Person (name, email or PeopleId)</label> '
      '<input type="text" name="person" value="{0}" size="28" placeholder="e.g. Smith or jane@example.com"> '
      '<label>Look back</label> <select name="gdays">{1}</select> '
      '<button class="ft-btn" type="submit">Check</button></form></div>'.format(
          h_escape(person),
          "".join('<option value="{0}"{1}>{0} days</option>'.format(
              d, " selected" if d == days else "") for d in (3, 7, 14, 30, 60))))

    if person:
        a('<h3 style="margin-top:18px">Person lookup</h3>')
        a(render_person_lookup(person, since, settings, days))
        a('<p style="font-size:12px"><a href="?tab=gates&amp;gdays={0}">Clear the person '
          'search</a></p>'.format(days))

    try:
        a(render_findings(settings, days))
    except Exception as e:
        a('<div class="ft-note">Could not run the problem checks: ' + h_escape(str(e)) + '</div>')

    # ---- the checks ---------------------------------------------------
    a('<h3 style="margin-top:18px">The checks, in the order TouchPoint runs them</h3>')
    a(render_gate_table(settings))
    if limit == 100:
        a('<p style="font-size:12px;color:#777">The IP check is off right now, so it cannot be '
          'the cause of a block that happens today. If a giver saw &ldquo;not supported from '
          'your current location&rdquo; while this reads 100, the setting was changed after '
          'the attempt. Ask TouchPoint support which value was in effect.</p>')

    # While a person is looked up, the site-wide tables are folded away so the
    # answer is not buried under them.
    if person:
        a('<details style="margin-top:22px"><summary style="cursor:pointer;font-size:15px">'
          '<b>All scored IP addresses and returning givers</b> (click to show)</summary>')

    # ---- scored IPs ----------------------------------------------------
    a('<h3 style="margin-top:22px">IP addresses scored in the last {0} days</h3>'.format(days))
    a('<p style="font-size:12px;color:#777">Click an IP to see who went through registration '
      'or a pay-balance link from it, and whether they paid.</p>')
    scored = list(q.QuerySql("""
        SELECT TOP 200 Evidence, Score, LastUpdated FROM dbo.FraudScores
        WHERE LastUpdated >= '{0}' ORDER BY LastUpdated DESC
    """.format(since.strftime("%Y-%m-%d"))))
    if not scored:
        a('<p>None. TouchPoint only scores an IP while the IP check is on, and only when '
          'someone reaches the payment step.</p>')
    else:
        ips = [str(r.Evidence) for r in scored]
        who = ip_people(ips, since)
        shared = shared_ip_counts(ips)
        nblocked = len([r for r in scored if ip_blocked(r.Score, limit)])
        a('<p>{0} IP(s) scored, <b>{1}</b> over the current limit. A score is saved the first '
          'time an IP reaches the payment step and refreshed after 30 days, so the time shown '
          'is the first attempt, not every one.</p>'.format(len(scored), nblocked))
        a('<div class="ft-scroll"><table class="ft-tbl"><thead><tr><th>IP</th><th>Score</th>'
          '<th>Scored</th><th class="num">Users this week</th><th>Signed in on this IP</th>'
          '</tr></thead><tbody>')
        for r in scored:
            ip = str(r.Evidence)
            users = shared.get(ip, 0)
            tag = (' <span class="muted">shared network</span>'
                   if users >= GATES_SHARED_IP_USERS else '')
            hot = ' style="background:#fdecea"' if ip_blocked(r.Score, limit) else ''
            a('<tr{0}><td>{1}{2}</td><td>{3}</td><td>{4}</td><td class="num">{5}</td>'
              '<td class="wrap">{6}</td></tr>'.format(
                  hot, _ip_link(ip, days), tag, _score_badge(r.Score, limit),
                  fmt_dt(r.LastUpdated), users, _names(who.get(ip, []))))
        a('</tbody></table></div>')
        a('<p style="font-size:11px;color:#777">&ldquo;Signed in on this IP&rdquo; is anyone '
          'who used TouchPoint from that address in the window. It shows who was there, not '
          'who tried to pay. A guest payment leaves no name.</p>')

    # ---- bypass log ---------------------------------------------------
    a('<h3 style="margin-top:22px">Givers who skipped the IP check</h3>')
    bypass = list(q.QuerySql("""
        SELECT TOP 25 al.ActivityDate, al.ClientIp, al.PeopleId, p.Name
        FROM dbo.ActivityLog al WITH (NOLOCK)
        LEFT JOIN dbo.People p ON p.PeopleId = al.PeopleId
        WHERE al.ActivityDate >= '{0}' AND al.Activity = '{1}'
        ORDER BY al.ActivityDate DESC
    """.format(since.strftime("%Y-%m-%d"), sql_escape(BYPASS_ACTIVITY))))
    a('<p style="font-size:13px">TouchPoint writes a line each time a signed-in giver with a '
      'recorded gift reaches the payment step. These people can never be blocked by the IP '
      'check. The line is written even while the check is off, so it does not prove the '
      'check was on.</p>')
    if not bypass:
        a('<p class="muted">None in this window.</p>')
    else:
        a('<table class="ft-tbl" style="width:auto"><thead><tr><th>When</th><th>Who</th>'
          '<th>IP</th></tr></thead><tbody>')
        for r in bypass:
            a('<tr><td>{0}</td><td><a href="/Person2/{1}" target="_blank">{2}</a></td>'
              '<td>{3}</td></tr>'.format(fmt_dt(r.ActivityDate), int(r.PeopleId or 0),
                                         h_escape(r.Name or ""), h_escape(r.ClientIp or "")))
        a('</tbody></table>')
    if person:
        a('</details>')

    a('<div class="ft-note" style="margin-top:22px"><b>Helping a blocked giver.</b> '
      'Ask them to sign in before paying if they have given before, or to switch networks '
      '(phone data instead of a VPN or office wifi). Staff can also enter the payment for '
      'them. If the blocked IP is your own building, every guest there is blocked until the '
      'score is refreshed; raise the limit or contact TouchPoint support.</div>')
    return "".join(out)


#####################################################################
#### MAIN
#####################################################################

def render_footer():
    """Small credit line. Kept quiet: this is a diagnostics page someone opens
    when something is already wrong, so it should not shout."""
    return (
        '<div style="margin:26px 0 8px;padding-top:12px;border-top:1px solid #e5e5e5;'
        'font-size:11px;color:#999;line-height:1.7">'
        'Fortis Troubleshooter &mdash; part of a free TouchPoint&reg; tool set by '
        '<a href="https://tpxisoftware.com" target="_blank" rel="noopener noreferrer" '
        'style="color:#777">TPxi Software</a>. '
        '<a href="https://github.com/bswaby/Touchpoint" target="_blank" '
        'rel="noopener noreferrer" style="color:#777">50+ more on GitHub</a>.'
        '<br>'
        'If these save your team time: '
        '<a href="https://displaycache.com" target="_blank" rel="noopener noreferrer" '
        'style="color:#777">DisplayCache</a>, church digital signage that integrates '
        'with TouchPoint&reg; &middot; '
        '<a href="https://tpxigo.com" target="_blank" rel="noopener noreferrer" '
        'style="color:#777">TPxi Go</a>, look up anyone and log calls and emails '
        'from Outlook or your phone.'
        '</div>')


def render_tabs(active):
    tabs = [("denials", "Denials"), ("health", "Health &amp; flow"),
            ("gates", "Blocked before Fortis")]
    out = ['<div style="border-bottom:2px solid #ddd;margin:14px 0 6px">']
    for key, label in tabs:
        if key == active:
            out.append('<span style="display:inline-block;padding:8px 18px;'
                       'border:2px solid #ddd;border-bottom:2px solid #fff;'
                       'margin-bottom:-2px;background:#fff;font-weight:bold;'
                       'border-radius:4px 4px 0 0;color:#2c5aa0">{0}</span>'.format(label))
        else:
            out.append('<a href="?tab={0}" style="display:inline-block;padding:8px 18px;'
                       'color:#666;text-decoration:none">{1}</a>'.format(key, label))
    out.append('</div>')
    return "".join(out)


def render_health():
    """The health tab: is anything flowing, and is it normal for the season."""
    out = []
    a = out.append

    methods = health_methods()
    if not methods:
        return ('<div class="ft-note">No rows in dbo.FortisWebhookTransactions '
                'in the last 12 months, so there is nothing to report on.</div>')

    hb = health_heartbeat()
    windows = health_window_counts(methods)
    assessment = health_assess(methods, hb, windows)
    current = [w for w in windows if w["is_current"]][0]
    baseline = [w for w in windows if not w["is_current"]]

    a('<div class="ft-note">'
      '<b>Why this tab exists.</b> The Denials tab reads an <b>error-only</b> log, so it '
      'can tell you why something failed but never that nothing is happening. A payment '
      'method that stops being submitted produces no errors at all. This tab reads the '
      'actual transaction stream instead, so silence is visible.'
      '</div>')

    # ---- heartbeat ---------------------------------------------------
    a('<h3 style="margin-top:18px">Is anything flowing?</h3>')
    a('<div class="ft-scroll"><table class="ft-tbl"><thead><tr>'
      '<th>Method</th><th>Status</th><th>Last transaction</th><th class="num">Age</th>'
      '<th class="num">Last {0}d</th><th class="num">Last 30d</th>'
      '<th style="width:40%">Assessment</th>'
      '</tr></thead><tbody>'.format(HEALTH_WINDOW_DAYS))
    for row in assessment:
        if row["status"] == "silent":
            badge = '<span style="background:#c0392b;color:#fff;padding:2px 8px;border-radius:3px">SILENT</span>'
        elif row["status"] == "quiet":
            badge = '<span style="background:#f0ad4e;color:#fff;padding:2px 8px;border-radius:3px">quiet</span>'
        else:
            badge = '<span style="background:#27ae60;color:#fff;padding:2px 8px;border-radius:3px">ok</span>'
        hrs = row["hours_ago"]
        age = ("%d h" % hrs) if hrs < 48 else ("%d days" % (hrs // 24))
        a('<tr><td><b>{0}</b></td><td>{1}</td>'
          '<td style="white-space:nowrap">{2}</td><td class="num">{3}</td>'
          '<td class="num">{4}</td><td class="num">{5}</td>'
          '<td class="wrap" style="font-size:12px">{6}</td></tr>'.format(
              h_escape(row["method"].upper()), badge,
              fmt_local(row["last_utc"]),
              age, row["now"], row["last30"], h_escape(row["why"])))
    a('</tbody></table></div>')
    a('<p style="font-size:11px;color:#777">Times shown are UTC{0:+d}h. Ages are computed '
      'in UTC, so the offset only affects how times read.</p>'.format(UTC_OFFSET_HOURS))

    # ---- baseline ----------------------------------------------------
    a('<h3 style="margin-top:22px">Compared with the same days in previous months</h3>')
    a('<div class="ft-note" style="font-size:12px">'
      'Giving is strongly monthly &mdash; volume spikes on the 1st and tails off &mdash; so '
      'week-on-week comparison flags every normal mid-month lull as an outage. These columns '
      'are the <b>same calendar days</b> in each month, which is a like-for-like test.'
      '</div>')
    # width:auto so the columns size to their contents. At 100% the numbers
    # stretch to the far right of each column and stop reading as a row.
    a('<div class="ft-scroll"><table class="ft-tbl" '
      'style="width:auto;min-width:520px">'
      '<thead><tr><th>Method</th>')
    for w in baseline:
        a('<th class="num" style="min-width:104px">{0}'
          '<div style="font-weight:normal;font-size:10px;color:#777;'
          'line-height:13px">{1}</div></th>'.format(
              h_escape(w["label"]), h_escape(w["range"])))
    a('<th class="num" style="background:#eef4fb;min-width:104px">now'
      '<div style="font-weight:normal;font-size:10px;color:#777;'
      'line-height:13px">{0}</div></th>'.format(h_escape(current["range"])))
    a('</tr></thead><tbody>')
    for row in assessment:
        m = row["method"]
        a('<tr><td><b>{0}</b></td>'.format(h_escape(m.upper())))
        for w in baseline:
            c = w["counts"].get(m, {"n": 0, "amt": 0.0})
            a('<td class="num">{0}<div style="font-size:10px;color:#777;'
              'line-height:13px">${1:,.0f}</div></td>'.format(c["n"], c["amt"]))
        c = current["counts"].get(m, {"n": 0, "amt": 0.0})
        style = ("background:#fdecea;font-weight:bold"
                 if row["status"] == "silent" else "background:#eef4fb")
        a('<td class="num" style="{0}">{1}<div style="font-size:10px;color:#777;'
          'font-weight:normal;line-height:13px">${2:,.0f}</div></td>'.format(
              style, c["n"], c["amt"]))
        a('</tr>')
    a('</tbody></table></div>')

    # ---- ACH settlement pipeline -------------------------------------
    a('<h3 style="margin-top:22px">ACH settlement pipeline</h3>')
    try:
        stuck = health_ach_pipeline()
    except Exception as e:
        stuck = []
        a('<div class="ft-note">Could not read the pipeline: ' + h_escape(str(e)) + '</div>')
    a('<div class="ft-note" style="font-size:12px">'
      'ACH moves <b>131 &rarr; 132 &rarr; 134</b>, reaching 134 with a settle date. '
      'Card authorises and captures in one step and never emits a settlement webhook, so it '
      'is not tracked here and its missing settle dates are not a fault.'
      '</div>')
    late = [r for r in stuck if int(r.AgeDays or 0) >= ACH_SETTLE_WARN_DAYS]
    if not stuck:
        a('<p style="color:#27ae60"><b>Nothing outstanding.</b> Every ACH transaction in the '
          'last 90 days reached settled status.</p>')
    else:
        a('<p>{0} ACH transaction(s) not yet settled, {1} of them older than {2} days.</p>'.format(
            len(stuck), len(late), ACH_SETTLE_WARN_DAYS))
        a('<div class="ft-scroll"><table class="ft-tbl"><thead><tr><th>Created</th><th>Who</th>'
          '<th class="num">Amount</th><th class="num">Status</th><th class="num">Age</th>'
          '</tr></thead><tbody>')
        for r in stuck[:40]:
            age = int(r.AgeDays or 0)
            warn = ' style="background:#fdecea"' if age >= ACH_SETTLE_WARN_DAYS else ''
            a('<tr{0}><td>{1}</td><td>{2}</td><td class="num">${3:,.2f}</td>'
              '<td class="num">{4}</td><td class="num">{5} d</td></tr>'.format(
                  warn,
                  fmt_local(r.CreatedUtc, ""),
                  h_escape(r.Who or ""), float(r.Amt or 0), r.MaxStatus, age))
        a('</tbody></table></div>')

    # ---- monthly trend ----------------------------------------------
    a('<h3 style="margin-top:22px">Twelve month trend</h3>')
    try:
        monthly = health_monthly()
        errors = health_monthly_errors()
    except Exception as e:
        monthly, errors = [], []
        a('<div class="ft-note">Could not build the trend: ' + h_escape(str(e)) + '</div>')

    err_map = {}
    for e in errors:
        err_map[(e.Mon, e.Kind)] = int(e.N or 0)

    by_month = {}
    for r in monthly:
        by_month.setdefault(r.Mon, {})[r.Method] = r

    a('<div class="ft-scroll"><table class="ft-tbl"><thead><tr>'
      '<th rowspan="2">Month</th>')
    for m in sorted(methods):
        a('<th class="num grp" colspan="4" style="text-align:center">{0}</th>'.format(
            h_escape(m.upper())))
    a('</tr><tr>')
    for m in sorted(methods):
        a('<th class="num grp" style="font-size:10px">count</th>'
          '<th class="num" style="font-size:10px">value</th>'
          '<th class="num" style="font-size:10px">approved</th>'
          '<th class="num" style="font-size:10px">denials</th>')
    a('</tr></thead><tbody>')
    for mon in sorted(by_month.keys()):
        a('<tr><td style="white-space:nowrap"><b>{0}</b></td>'.format(h_escape(mon)))
        for m in sorted(methods):
            r = by_month[mon].get(m)
            if not r:
                a('<td class="num grp">0</td><td class="num">-</td>'
                  '<td class="num">-</td><td class="num">-</td>')
                continue
            n = int(r.N or 0)
            appr = int(r.Approved or 0)
            errs = err_map.get((mon, m), 0)
            a('<td class="num grp">{0}</td><td class="num">${1:,.0f}</td>'
              '<td class="num">{2}</td><td class="num">{3}</td>'.format(
                  n, float(r.Amt or 0),
                  ("%.0f%%" % (100.0 * appr / n)) if n else "-",
                  ("%.1f%%" % (100.0 * errs / n)) if n else "-"))
        a('</tr>')
    a('</tbody></table></div>')
    a('<p style="font-size:11px;color:#777">"errors" is denials from FortisTransactionLog as a '
      'percentage of that method\'s transactions, not a raw count. A count rises purely with '
      'volume; a rate is comparable month to month.</p>')

    # ---- alert status ------------------------------------------------
    st = health_load_state()
    a('<h3 style="margin-top:22px">Alerting</h3>')
    if HEALTH_ALERT_PEOPLE_ID:
        a('<p>Emailing PeopleId <b>{0}</b> when a method goes silent, at most once every '
          '{1} hours, plus once when it recovers.<br>'
          '<span style="font-size:12px;color:#777">Last scheduled check: {2}</span></p>'.format(
              HEALTH_ALERT_PEOPLE_ID, HEALTH_ALERT_COOLDOWN_HOURS,
              h_escape(st.get("last_check") or "never run")))
    else:
        a('<div class="ft-note"><b>Email alerting is off.</b> Set '
          '<code>HEALTH_ALERT_PEOPLE_ID</code> to a PeopleId, then add this to Special '
          'Content &gt; Python &gt; ScheduledTasks so it is checked without anyone '
          'remembering to look:<br>'
          '<pre style="font-size:11px">'
          '# &gt;&gt;&gt; Fortis health check &gt;&gt;&gt;\n'
          'try:\n'
          '    Data.fortis_health = \'true\'\n'
          '    model.CallScript(\'FortisTroubleshooter\')\n'
          'except Exception as _fh_e:\n'
          '    print \'Fortis health error: \' + str(_fh_e)\n'
          '# &lt;&lt;&lt; Fortis health check &lt;&lt;&lt;'
          '</pre></div>')
    return "".join(out)


def main():
    model.Header = PAGE_TITLE

    if not user_allowed():
        model.Form = ('<div class="ft-wrap">' + render_styles() +
                      '<h2>' + PAGE_TITLE + '</h2>'
                      '<div class="ft-note">You do not have permission to view this page. '
                      'Required role (one of): ' + ", ".join(ALLOWED_ROLES) + '.</div></div>')
        return

    tab = (get_param("tab") or "denials").lower()

    if tab in ("health", "gates"):
        html = ['<div class="ft-wrap">', render_styles(),
                '<h2>' + PAGE_TITLE + '</h2>', render_tabs(tab)]
        try:
            html.append(render_health() if tab == "health" else render_gates())
        except Exception as e:
            import traceback
            html.append('<div class="ft-note"><b>This tab hit an error.</b>'
                        '<pre style="white-space:pre-wrap;font-size:11px;">'
                        + h_escape(str(e)) + '\n\n'
                        + h_escape(traceback.format_exc()) + '</pre>'
                        'The Denials tab is unaffected.</div>')
        html.append(render_footer())
        html.append('</div>')
        model.Form = "".join(html)
        return

    start_dt, end_dt = resolve_window()
    category = get_param("category")
    search = get_param("search")

    # base querystring to preserve date/search when clicking cards
    base_qs = "start={0}&end={1}".format(
        start_dt.strftime("%Y-%m-%d"),
        (end_dt - datetime.timedelta(days=1)).strftime("%Y-%m-%d"))
    if search:
        base_qs += "&search=" + search

    html = ['<div class="ft-wrap">']
    html.append(render_styles())
    html.append('<h2>' + PAGE_TITLE + '</h2>')
    html.append(render_tabs(tab))
    html.append(
        '<div class="ft-note">'
        '<b>What am I looking at?</b> Every row below is a request the Fortis gateway '
        '<b>rejected</b> (from dbo.FortisTransactionLog). One thing worth knowing before you '
        'send examples to Fortis &mdash; the <b>Last 4</b> is only present when the rejected '
        'request actually carried a card:<br>'
        '&bull; <b>CC Sale</b> and <b>ACH Debit</b> requests carry the masked account '
        '(e.g. "************0093"), so <b>Last 4 is shown</b>.<br>'
        '&bull; <b>Contact Sync</b> requests (/v1/contacts) carry name + address only and '
        '<b>no card at all</b>, so there is <b>no Last 4</b> for those &mdash; the card never '
        'entered the request. Street-length denials show up on <b>both</b> call types, so some '
        'street rows have a Last 4 and some (the contact-sync ones) do not.<br>'
        'Full card numbers are never stored; only the last 4 is ever available.'
        '</div>')

    try:
        rows = get_rows(start_dt, end_dt, search)
    except Exception as e:
        html.append('<div class="ft-note">Error reading FortisTransactionLog: ' +
                    h_escape(str(e)) + '</div></div>')
        model.Form = "".join(html)
        return

    # Parse all rows once
    all_records = []
    for row in rows:
        try:
            all_records.append(build_record(row))
        except:
            continue

    # Summary counts (across the whole window/search, before category filter)
    counts = {}
    for rec in all_records:
        counts[rec["category"]] = counts.get(rec["category"], 0) + 1

    # Detail set (apply category filter)
    if category:
        detail_records = [r for r in all_records if r["category"] == category]
    else:
        detail_records = all_records

    html.append(render_filters(start_dt, end_dt, category, search, None))
    try:
        denoms = denial_denominators(start_dt, end_dt)
    except Exception:
        denoms = {}
    html.append(render_cards(counts, category, base_qs, denoms))
    try:
        html.append(render_rate_trend(denial_rate_trend()))
    except Exception:
        pass
    html.append('<p style="font-size:12px;color:#666;">Showing <b>{0}</b> of {1} denial(s) in range{2}.</p>'.format(
        len(detail_records), len(all_records),
        (" matching category '%s'" % CATEGORY_LABELS.get(category, category)) if category else ""))
    html.append(render_actions_js())
    html.append(render_table(detail_records))
    html.append(render_footer())
    html.append('</div>')

    model.Form = "".join(html)


def _called_from_scheduler():
    for flag in ("fortis_health", "scheduler"):
        try:
            if str(getattr(model.Data, flag, "")).lower() == "true":
                return True
        except Exception:
            pass
    return False


try:
    if _called_from_scheduler():
        # No page, no output. health_run_scheduled swallows everything: an
        # exception escaping here would make TouchPoint email the admins on
        # every scheduled run.
        health_run_scheduled()
    else:
        main()
except Exception as _fatal:
    import traceback
    model.Header = PAGE_TITLE
    model.Form = (
        '<div style="font-family:Segoe UI,Arial,sans-serif;">'
        '<h2>Fortis Denial Troubleshooter</h2>'
        '<div style="background:#fff0f0;border-left:4px solid #c0392b;padding:12px 16px;">'
        '<b>The page hit an error.</b><br><pre style="white-space:pre-wrap;font-size:12px;">'
        + str(_fatal) + '\n\n' + traceback.format_exc() +
        '</pre></div></div>'
    )
