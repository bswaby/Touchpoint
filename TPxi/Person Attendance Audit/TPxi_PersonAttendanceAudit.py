#roles=Admin
# -*- coding: utf-8 -*-
#####################################################################
# TPxi_PersonAttendanceAudit  --  READ-ONLY safe-church audit
#####################################################################
# Person-centric co-attendance audit. Point it at ONE person and it
# reconstructs, from actual attendance (AttendanceFlag=1):

#####################################################################

import re
import datetime
import traceback

model.Header = "Person Attendance Audit (read-only)"

MINOR_AGE = 18    # under 18 at time of attendance = minor (matches ACS script)
CAP = 40000       # safety cap on co-attendance rows


def _param(name, default=''):
    try:
        v = getattr(model.Data, name)
        v = str(v).strip()
        return v if v else default
    except:
        return default


def _isodate(s):
    return s if re.match(r'^\d{4}-\d{2}-\d{2}$', s or '') else ''


PID = _param('pid', '')
QSEARCH = _param('q', '')
CO = _param('co', '')
try:
    PID = int(PID) if PID else 0
except:
    PID = 0
try:
    CO = int(CO) if CO else 0
except:
    CO = 0

DAYS = _param('days', '')
try:
    DAYS = int(DAYS) if DAYS else 0
except:
    DAYS = 0
FROM = _isodate(_param('from', ''))
TO = _isodate(_param('to', ''))

if FROM or TO:
    WINDOW_LABEL = "%s to %s" % (FROM or 'start', TO or 'now')
elif DAYS:
    WINDOW_LABEL = "last %d days" % DAYS
else:
    WINDOW_LABEL = "all history"


def date_pred(col):
    if FROM and TO:
        return "%s >= '%s' AND %s < DATEADD(DAY, 1, '%s')" % (col, FROM, col, TO)
    if FROM:
        return "%s >= '%s'" % (col, FROM)
    if TO:
        return "%s < DATEADD(DAY, 1, '%s')" % (col, TO)
    if DAYS:
        return "%s >= DATEADD(DAY, -%d, GETDATE())" % (col, DAYS)
    return "1 = 1"


# ---- unicode / html helpers (ASCII-safe; preserves chars as entities) -----
def _u(val):
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


_WD = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']


def weekday(iso):
    try:
        y, m, d = int(iso[0:4]), int(iso[5:7]), int(iso[8:10])
        return _WD[datetime.date(y, m, d).weekday()]
    except:
        return ''


def _d(iso):
    try:
        return datetime.date(int(iso[0:4]), int(iso[5:7]), int(iso[8:10]))
    except:
        return None


def days_between(a, b):
    da, db = _d(a), _d(b)
    if da is None or db is None:
        return None
    return abs((db - da).days)


def dur_label(days):
    # friendly span; not thresholded -- just how long first->last co-attendance was
    if days is None:
        return ''
    if days <= 0:
        return 'same day'
    yrs = days // 365
    mo = int(round((days - yrs * 365) / 30.4))
    if mo >= 12:
        yrs += 1
        mo = 0
    if yrs >= 1:
        return "%dy %dmo" % (yrs, mo) if mo else "%dy" % yrs
    if days >= 14:
        return "%dwk" % int(round(days / 7.0))
    return "%dd" % days


# ---------------------------------------------------------------------------
# DATA
# ---------------------------------------------------------------------------
def fetch_target(pid):
    sql = """
    SELECT p.PeopleId, p.Name2 AS Name, p.FamilyId,
           CONVERT(varchar(10), p.BDate, 126) AS DOB,
           DATEDIFF(YEAR, p.BDate, GETDATE())
             - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, p.BDate, GETDATE()), p.BDate) > GETDATE() THEN 1 ELSE 0 END AS AgeNow,
           ISNULL(g.Description, '') AS Gender,
           ISNULL(ms.Description, '') AS MemberStatus,
           ISNULL(p.IsDeceased, 0) AS IsDeceased, ISNULL(p.ArchivedFlag, 0) AS ArchivedFlag
    FROM dbo.People p
    LEFT JOIN lookup.Gender g ON g.Id = p.GenderId
    LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
    WHERE p.PeopleId = %d
    """ % pid
    for r in q.QuerySql(sql):
        return {'PeopleId': r.PeopleId, 'Name': _u(r.Name), 'FamilyId': r.FamilyId or 0,
                'DOB': r.DOB or '', 'AgeNow': r.AgeNow, 'Gender': _u(r.Gender),
                'MemberStatus': _u(r.MemberStatus),
                'Deceased': (int(r.IsDeceased or 0) == 1), 'Archived': (int(r.ArchivedFlag or 0) == 1)}
    return None


def fetch_sessions(pid):
    # every meeting the target was present for, with their role in that org
    sql = """
    SELECT m.MeetingId, m.OrganizationId, ISNULL(o.OrganizationName, '') AS OrganizationName,
           CONVERT(varchar(10), m.MeetingDate, 126) AS MeetingDate,
           ISNULL(mt.Description, '') AS TargetRole
    FROM dbo.Attend a
    JOIN dbo.Meetings m ON m.MeetingId = a.MeetingId AND m.DidNotMeet = 0
    JOIN dbo.Organizations o ON o.OrganizationId = m.OrganizationId
    LEFT JOIN dbo.OrganizationMembers om ON om.OrganizationId = m.OrganizationId AND om.PeopleId = %d
    LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
    WHERE a.PeopleId = %d AND a.AttendanceFlag = 1 AND %s
    """ % (pid, pid, date_pred('m.MeetingDate'))
    out = []
    for r in q.QuerySql(sql):
        out.append({'MeetingId': r.MeetingId, 'OrgId': r.OrganizationId,
                    'Org': _u(r.OrganizationName), 'Date': r.MeetingDate or '',
                    'Role': _u(r.TargetRole)})
    return out


def fetch_co(pid, fam):
    sql = """
    SET NOCOUNT ON;
    IF OBJECT_ID('tempdb..#tm') IS NOT NULL DROP TABLE #tm;
    SELECT DISTINCT a.MeetingId, m.OrganizationId, m.MeetingDate
    INTO #tm
    FROM dbo.Attend a
    JOIN dbo.Meetings m ON m.MeetingId = a.MeetingId AND m.DidNotMeet = 0
    WHERE a.PeopleId = %d AND a.AttendanceFlag = 1 AND %s;

    SELECT TOP %d
        t.MeetingId, t.OrganizationId, ISNULL(o.OrganizationName, '') AS OrganizationName,
        CONVERT(varchar(10), t.MeetingDate, 126) AS MeetingDate,
        a2.PeopleId, p.Name2 AS Name, CONVERT(varchar(10), p.BDate, 126) AS DOB,
        ISNULL(g.Description, '') AS Gender, ISNULL(p.FamilyId, 0) AS FamilyId,
        ISNULL(p.IsDeceased, 0) AS IsDeceased, ISNULL(p.ArchivedFlag, 0) AS ArchivedFlag,
        ISNULL(ms.Description, '') AS MemberStatus,
        CASE WHEN p.BDate IS NULL THEN NULL ELSE
             DATEDIFF(YEAR, p.BDate, t.MeetingDate)
               - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, p.BDate, t.MeetingDate), p.BDate) > t.MeetingDate THEN 1 ELSE 0 END
        END AS AgeAtAttend,
        ISNULL(mt.Description, '') AS RoleInGroup,
        CASE WHEN ISNULL(p.FamilyId, 0) = %d AND %d <> 0 THEN 1 ELSE 0 END AS SameFamily
    FROM #tm t
    JOIN dbo.Attend a2 ON a2.MeetingId = t.MeetingId AND a2.AttendanceFlag = 1 AND a2.PeopleId <> %d
    JOIN dbo.Organizations o ON o.OrganizationId = t.OrganizationId
    JOIN dbo.People p ON p.PeopleId = a2.PeopleId
    LEFT JOIN lookup.Gender g ON g.Id = p.GenderId
    LEFT JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId
    LEFT JOIN dbo.OrganizationMembers om ON om.OrganizationId = t.OrganizationId AND om.PeopleId = a2.PeopleId
    LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
    ORDER BY t.MeetingDate DESC, p.Name2;
    """ % (pid, date_pred('m.MeetingDate'), CAP, fam, fam, pid)
    out = []
    for r in q.QuerySql(sql):
        age = None if r.AgeAtAttend is None else int(r.AgeAtAttend)
        out.append({'MeetingId': r.MeetingId, 'OrgId': r.OrganizationId, 'Org': _u(r.OrganizationName),
                    'Date': r.MeetingDate or '', 'PeopleId': r.PeopleId, 'Name': _u(r.Name),
                    'DOB': r.DOB or '', 'Gender': _u(r.Gender), 'FamilyId': r.FamilyId or 0,
                    'Age': age, 'Minor': (age is not None and age < MINOR_AGE),
                    'Role': _u(r.RoleInGroup), 'SameFam': (int(r.SameFamily or 0) == 1),
                    'Deceased': (int(r.IsDeceased or 0) == 1), 'Archived': (int(r.ArchivedFlag or 0) == 1),
                    'MemberStatus': _u(r.MemberStatus)})
    return out


def fetch_roles(pid):
    # every involvement the subject is/was a member of, with role + status
    sql = """
    SELECT o.OrganizationId, ISNULL(o.OrganizationName, '') AS Org,
           ISNULL(mt.Description, '') AS Role,
           ISNULL(os.Description, '') AS OrgStatus,
           CONVERT(varchar(10), om.EnrollmentDate, 126) AS Enrolled,
           CONVERT(varchar(10), om.InactiveDate, 126) AS Inactive
    FROM dbo.OrganizationMembers om
    JOIN dbo.Organizations o ON o.OrganizationId = om.OrganizationId
    LEFT JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId
    LEFT JOIN lookup.OrganizationStatus os ON os.Id = o.OrganizationStatusId
    WHERE om.PeopleId = %d
    ORDER BY CASE WHEN om.InactiveDate IS NULL THEN 0 ELSE 1 END, o.OrganizationName
    """ % pid
    out = []
    for r in q.QuerySql(sql):
        role = _u(r.Role)
        out.append({'OrgId': r.OrganizationId, 'Org': _u(r.Org), 'Role': role,
                    'OrgStatus': _u(r.OrgStatus), 'Enrolled': r.Enrolled or '', 'Inactive': r.Inactive or '',
                    'IsLeader': any(w in role.lower() for w in ('lead', 'teach', 'volunt', 'coord', 'director'))})
    return out


def fetch_background(pid):
    # clearance summary (Volunteer) + background-check history
    vol = None
    for r in q.QuerySql("""
        SELECT ISNULL(vs.Description,'(none)') AS StatusDesc,
               CONVERT(varchar(10), v.ProcessedDate, 126) AS Processed,
               ISNULL(v.Standard,0) AS Standard, ISNULL(v.Children,0) AS Children, ISNULL(v.Leader,0) AS Leader,
               CONVERT(varchar(10), v.MVRProcessedDate, 126) AS MVR,
               CONVERT(varchar(10), v.TrainingDate, 126) AS Trained
        FROM dbo.Volunteer v LEFT JOIN lookup.VolApplicationStatus vs ON vs.Id = v.StatusId
        WHERE v.PeopleId = %d""" % pid):
        vol = {'Status': _u(r.StatusDesc), 'Processed': r.Processed or '',
               'Standard': int(r.Standard or 0), 'Children': int(r.Children or 0), 'Leader': int(r.Leader or 0),
               'MVR': r.MVR or '', 'Trained': r.Trained or ''}
    checks = []
    for r in q.QuerySql("""
        SELECT CONVERT(varchar(10), Created, 126) AS Created,
               ISNULL(CONVERT(varchar(50), ApprovalStatus), '') AS Approval,
               ISNULL(CONVERT(varchar(50), [Level]), '') AS Lvl,
               ISNULL(ChildServing, 0) AS ChildServing
        FROM dbo.BackgroundChecks WHERE PeopleID = %d ORDER BY Created DESC""" % pid):
        checks.append({'Created': r.Created or '', 'Approval': _u(r.Approval), 'Level': _u(r.Lvl),
                       'ChildServing': int(r.ChildServing or 0)})
    return vol, checks


def fetch_documents(pid):
    # ALL documents/forms on file for the person (VolunteerForm), no name filter
    sql = """
    SELECT vf.Id, ISNULL(vf.Name, '') AS Name, CONVERT(varchar(10), vf.AppDate, 126) AS AppDate,
           ISNULL(vf.IsDocument, 0) AS IsDoc, ISNULL(vf.LargeId, 0) AS LargeId,
           ISNULL(up.Name2, '') AS Uploader
    FROM dbo.VolunteerForm vf
    LEFT JOIN dbo.People up ON up.PeopleId = vf.UploaderId
    WHERE vf.PeopleId = %d
    ORDER BY vf.AppDate DESC, vf.Id DESC
    """ % pid
    out = []
    for r in q.QuerySql(sql):
        out.append({'Id': r.Id, 'Name': _u(r.Name), 'AppDate': r.AppDate or '',
                    'IsDoc': int(r.IsDoc or 0), 'LargeId': r.LargeId or 0, 'Uploader': _u(r.Uploader)})
    return out


def fetch_access(pid):
    # TouchPoint login(s) + roles (system access the subject had/has)
    sql = """
    SELECT u.UserId, ISNULL(u.Username, '') AS Username,
           CONVERT(varchar(10), u.LastLoginDate, 126) AS LastLogin,
           CONVERT(varchar(10), u.CreationDate, 126) AS Created,
           ISNULL(u.IsLockedOut, 0) AS Locked,
           (SELECT STRING_AGG(r.RoleName, ', ') FROM dbo.UserRole ur
            JOIN dbo.Roles r ON r.RoleId = ur.RoleId WHERE ur.UserId = u.UserId) AS Roles
    FROM dbo.Users u WHERE u.PeopleId = %d
    ORDER BY u.LastLoginDate DESC
    """ % pid
    out = []
    for r in q.QuerySql(sql):
        out.append({'Username': _u(r.Username), 'LastLogin': r.LastLogin or '', 'Created': r.Created or '',
                    'Locked': (int(r.Locked or 0) == 1), 'Roles': _u(r.Roles)})
    return out


def fetch_comms(pid):
    # in-system messaging the subject SENT (email + text blasts) and inbound
    # texts FROM them. Recipient + minor-recipient counts let you separate a
    # personal 1:1 from a bulk broadcast. NOTE: mostly bulk; off-platform
    # (personal cell/email/DM) is NOT captured -- see Method sheet caveat.
    m = MINOR_AGE
    sql = """
    SET NOCOUNT ON;
    SELECT TOP 3000 * FROM (
      SELECT 'Email' AS Channel, CONVERT(varchar(10), eq.Sent, 126) AS Dt,
             LEFT(ISNULL(eq.Subject, ''), 160) AS Summary,
             COUNT(et.PeopleId) AS Recipients,
             SUM(CASE WHEN p.BDate IS NOT NULL AND (DATEDIFF(YEAR, p.BDate, eq.Sent)
                   - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, p.BDate, eq.Sent), p.BDate) > eq.Sent THEN 1 ELSE 0 END) < %d
                 THEN 1 ELSE 0 END) AS MinorRecipients
      FROM dbo.EmailQueue eq WITH (NOLOCK)
      JOIN dbo.EmailQueueTo et WITH (NOLOCK) ON et.Id = eq.Id
      JOIN dbo.People p WITH (NOLOCK) ON p.PeopleId = et.PeopleId
      WHERE eq.QueuedBy = %d AND eq.Sent IS NOT NULL AND %s
      GROUP BY eq.Id, eq.Sent, eq.Subject

      UNION ALL

      SELECT 'Text' AS Channel, CONVERT(varchar(10), sl.Created, 126) AS Dt,
             LEFT(ISNULL(sl.Message, ''), 160) AS Summary,
             COUNT(si.PeopleID) AS Recipients,
             SUM(CASE WHEN p.BDate IS NOT NULL AND (DATEDIFF(YEAR, p.BDate, sl.Created)
                   - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, p.BDate, sl.Created), p.BDate) > sl.Created THEN 1 ELSE 0 END) < %d
                 THEN 1 ELSE 0 END) AS MinorRecipients
      FROM dbo.SMSList sl WITH (NOLOCK)
      JOIN dbo.SMSItems si WITH (NOLOCK) ON si.ListID = sl.ID
      JOIN dbo.People p WITH (NOLOCK) ON p.PeopleId = si.PeopleID
      WHERE sl.SenderID = %d AND %s
      GROUP BY sl.ID, sl.Created, sl.Message

      UNION ALL

      SELECT 'Text (received)' AS Channel, CONVERT(varchar(10), sr.DateReceived, 126) AS Dt,
             LEFT(ISNULL(sr.Body, ''), 160) AS Summary, 1 AS Recipients, 0 AS MinorRecipients
      FROM dbo.SmsReceived sr WITH (NOLOCK)
      WHERE sr.FromPeopleId = %d AND %s
    ) z
    ORDER BY Dt DESC;
    """ % (m, pid, date_pred('eq.Sent'), m, pid, date_pred('sl.Created'), pid, date_pred('sr.DateReceived'))
    out = []
    for r in q.QuerySql(sql):
        rec = int(r.Recipients or 0)
        typ = '1:1' if rec == 1 else ('small (%d)' % rec if rec <= 5 else 'bulk (%d)' % rec)
        out.append({'Channel': _u(r.Channel), 'Date': r.Dt or '', 'Summary': _u(r.Summary),
                    'Recipients': rec, 'MinorRecipients': int(r.MinorRecipients or 0), 'Type': typ})
    return out


def fetch_checkouts(pid):
    # everyone the SUBJECT checked out (picked up) -- age-at-checkout flags minors
    sql = """
    SELECT p.PeopleId, p.Name2 AS Name, ISNULL(o.OrganizationName, '') AS Org,
           CONVERT(varchar(10), m.MeetingDate, 126) AS Dt,
           CONVERT(varchar(10), p.BDate, 126) AS DOB, ISNULL(p.FamilyId,0) AS FamilyId,
           CASE WHEN p.BDate IS NULL THEN NULL ELSE
                DATEDIFF(YEAR, p.BDate, m.MeetingDate)
                  - CASE WHEN DATEADD(YEAR, DATEDIFF(YEAR, p.BDate, m.MeetingDate), p.BDate) > m.MeetingDate THEN 1 ELSE 0 END
           END AS AgeAtCheckout
    FROM dbo.Attend a
    JOIN dbo.Meetings m ON m.MeetingId = a.MeetingId AND m.DidNotMeet = 0
    JOIN dbo.People p ON p.PeopleId = a.PeopleId
    JOIN dbo.Organizations o ON o.OrganizationId = m.OrganizationId
    WHERE a.CheckOutBy = %d AND a.PeopleId <> %d AND %s
    ORDER BY m.MeetingDate DESC, p.Name2
    """ % (pid, pid, date_pred('m.MeetingDate'))
    out = []
    for r in q.QuerySql(sql):
        age = None if r.AgeAtCheckout is None else int(r.AgeAtCheckout)
        out.append({'PeopleId': r.PeopleId, 'Name': _u(r.Name), 'Org': _u(r.Org), 'Date': r.Dt or '',
                    'DOB': r.DOB or '', 'FamilyId': r.FamilyId or 0, 'Age': age,
                    'Minor': (age is not None and age < MINOR_AGE)})
    return out


# ---------------------------------------------------------------------------
# STYLE
# ---------------------------------------------------------------------------
CSS = """
<style>
  .paa { font-family:'Segoe UI',-apple-system,Arial,sans-serif; color:#1e293b; max-width:1200px; line-height:1.5; }
  .paa h2 { color:#7f1d1d; margin:0 0 2px; font-size:24px; }
  .paa .lede { color:#64748b; font-size:13px; margin:0 0 14px; max-width:940px; }
  .paa code { background:#eef2f7; padding:1px 5px; border-radius:3px; font-size:12px; }
  .paa .subject { background:#fef2f2; border:1px solid #fecaca; border-left:5px solid #b91c1c; border-radius:10px; padding:12px 16px; margin:0 0 14px; }
  .paa .subject .nm { font-size:20px; font-weight:800; color:#7f1d1d; }
  .paa .subject .meta { font-size:13px; color:#7a4a4a; margin-top:2px; }
  .paa .cards { display:flex; gap:12px; flex-wrap:wrap; margin:0 0 16px; }
  .paa .card { flex:1; min-width:150px; border:1px solid #e2e8f0; border-top:3px solid #7f1d1d; border-radius:10px; padding:12px 14px; background:#fff; }
  .paa .card .lbl { font-size:11px; text-transform:uppercase; letter-spacing:.6px; color:#64748b; font-weight:600; }
  .paa .card .val { font-size:26px; font-weight:800; margin-top:3px; color:#7f1d1d; }
  .paa .card .sub { font-size:11px; color:#94a3b8; }
  .paa .card.minor { border-top-color:#be123c; } .paa .card.minor .val { color:#be123c; }
  .paa .filters { display:flex; flex-wrap:wrap; align-items:flex-end; gap:14px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px 16px; margin:0 0 14px; }
  .paa .filters label { font-size:11px; text-transform:uppercase; letter-spacing:.5px; color:#64748b; font-weight:600; display:flex; flex-direction:column; gap:3px; }
  .paa .filters input { font-size:14px; padding:6px 9px; border:1px solid #cbd5e1; border-radius:6px; }
  .paa .filters .go { background:#7f1d1d; color:#fff; border:0; border-radius:6px; padding:8px 18px; font-size:14px; cursor:pointer; }
  .paa .filters .reset { color:#64748b; font-size:12px; align-self:center; }
  .paa .sec { margin:22px 0 6px; font-size:17px; color:#7f1d1d; font-weight:800; border-bottom:2px solid #7f1d1d; padding-bottom:3px; }
  .paa .sec .c { color:#94a3b8; font-size:13px; font-weight:500; }
  .paa .toolbar { display:flex; align-items:center; gap:14px; flex-wrap:wrap; margin:0 0 8px; }
  .paa .btn { display:inline-block; background:#7f1d1d; color:#fff; padding:7px 14px; border-radius:6px; font-size:13px; border:0; cursor:pointer; }
  .paa .btn.alt { background:#0f766e; }
  .paa .toolbar .chk { display:flex; align-items:center; gap:6px; font-size:13px; color:#334155; }
  .paa .toolbar .count { color:#64748b; font-size:13px; margin-left:auto; }
  .paa table { border-collapse:collapse; width:100%; font-size:13px; margin:2px 0 6px; }
  .paa th { background:#eef2f7; color:#334155; text-align:left; padding:5px 9px; font-size:11px; text-transform:uppercase; letter-spacing:.4px; cursor:pointer; white-space:nowrap; }
  .paa th:hover { background:#e2e8f0; }
  .paa th .arrow { color:#94a3b8; font-weight:400; }
  .paa thead .flt th { background:#fff; padding:3px 6px; cursor:auto; }
  .paa thead .flt input { width:100%; box-sizing:border-box; font-size:12px; padding:3px 5px; border:1px solid #cbd5e1; border-radius:4px; font-weight:400; text-transform:none; letter-spacing:0; }
  .paa td { padding:5px 9px; border-bottom:1px solid #f1f5f9; }
  .paa a { color:#7f1d1d; text-decoration:none; } .paa a:hover { text-decoration:underline; }
  .paa .num { text-align:center; font-variant-numeric:tabular-nums; }
  .paa .minorflag { color:#be123c; font-weight:800; }
  .paa .fam { color:#b45309; font-size:11px; font-weight:700; }
  .paa .muted { color:#94a3b8; font-size:12px; }
  .paa .cap { background:#fff7ed; border:1px solid #fed7aa; border-left:4px solid #ea580c; color:#7a4a00; border-radius:8px; padding:8px 14px; margin:0 0 12px; font-size:12.5px; }
  .paa .search { display:flex; gap:10px; align-items:flex-end; }
  .paa .rlist a { display:block; padding:6px 4px; border-bottom:1px solid #f1f5f9; }
  .paa .tabs { display:flex; flex-wrap:wrap; gap:4px; border-bottom:2px solid #7f1d1d; margin:16px 0 12px; }
  .paa .tabs .tab { padding:7px 14px; font-size:13px; font-weight:700; color:#7f1d1d; cursor:pointer; border:1px solid #e2e8f0; border-bottom:none; border-radius:6px 6px 0 0; background:#f8fafc; }
  .paa .tabs .tab:hover { background:#eef2f7; }
  .paa .tabs .tab.active { background:#7f1d1d; color:#fff; border-color:#7f1d1d; }
  .paa .tabs .tab.allt { margin-left:auto; background:#0f766e; color:#fff; border-color:#0f766e; }
</style>
"""

PRINT_JS = """
<script>
function paaFilter(id){
  var t=document.getElementById(id); if(!t) return;
  var ins=t.querySelectorAll('thead tr.flt input'); var vals=[];
  for(var i=0;i<ins.length;i++) vals.push(ins[i].value.toLowerCase());
  var hb=document.getElementById('hide_'+id); var hide=hb&&hb.checked;
  var rows=t.tBodies[0].rows, shown=0;
  for(var r=0;r<rows.length;r++){
    var td=rows[r].cells, ok=true;
    for(var c=0;c<vals.length;c++){ if(vals[c] && (td[c].textContent||'').toLowerCase().indexOf(vals[c])<0){ok=false;break;} }
    if(ok && hide && rows[r].getAttribute('data-minor')!=='1') ok=false;
    rows[r].style.display=ok?'':'none'; if(ok) shown++;
  }
  var s=document.getElementById('shown_'+id); if(s) s.textContent=shown;
}
function paaXe(s){ s=(s==null?'':(''+s)); return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function paaXls(){
  var tables=document.querySelectorAll('table[data-sheet]');
  var xml='<?xml version="1.0"?>\\r\\n<?mso-application progid="Excel.Sheet"?>\\r\\n';
  xml+='<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\\r\\n';
  for(var ti=0;ti<tables.length;ti++){
    var t=tables[ti]; var nm=t.getAttribute('data-sheet');
    var ths=t.tHead?t.tHead.querySelectorAll('tr.hdr th'):[];
    var head=[]; for(var h=0;h<ths.length;h++) head.push((ths[h].textContent||'').replace(/[\\u2191\\u2193\\u21C5]/g,'').trim());
    var body=t.tBodies[0]?t.tBodies[0].rows:[];
    var hasIds = body.length && body[0].getAttribute('data-pid')!==null;
    if(hasIds) head=head.concat(['PeopleId','OrgId']);
    xml+='<Worksheet ss:Name="'+paaXe(nm)+'"><Table>';
    xml+='<Row>'; for(var c=0;c<head.length;c++) xml+='<Cell><Data ss:Type="String">'+paaXe(head[c])+'</Data></Cell>'; xml+='</Row>';
    for(var r=0;r<body.length;r++){
      var td=body[r].cells; xml+='<Row>';
      for(var c2=0;c2<td.length;c2++) xml+='<Cell><Data ss:Type="String">'+paaXe((td[c2].textContent||'').trim())+'</Data></Cell>';
      if(hasIds){ xml+='<Cell><Data ss:Type="String">'+paaXe(body[r].getAttribute('data-pid')||'')+'</Data></Cell>';
                  xml+='<Cell><Data ss:Type="String">'+paaXe(body[r].getAttribute('data-oid')||'')+'</Data></Cell>'; }
      xml+='</Row>';
    }
    xml+='</Table></Worksheet>\\r\\n';
  }
  xml+='</Workbook>';
  var blob=new Blob([xml],{type:'application/vnd.ms-excel'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='person_audit.xls';
  document.body.appendChild(a); a.click(); setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(a.href);},100);
}
var paaDir={};
function paaSort(id,c){
  var t=document.getElementById(id), tb=t.tBodies[0], rows=[].slice.call(tb.rows);
  var dir=paaDir[id+'_'+c]=-(paaDir[id+'_'+c]||1);
  rows.sort(function(a,b){var x=(a.cells[c].textContent||'').trim().toLowerCase(),y=(b.cells[c].textContent||'').trim().toLowerCase();
    var nx=parseFloat(x),ny=parseFloat(y);
    if(!isNaN(nx)&&!isNaN(ny)){return (nx-ny)*dir;}
    if(x<y)return -1*dir; if(x>y)return 1*dir; return 0;});
  for(var i=0;i<rows.length;i++) tb.appendChild(rows[i]);
}
function paaCsvCell(v){ v=(v==null?'':(''+v)); return '"'+v.replace(/"/g,'""')+'"'; }
function paaExport(id, fname){
  var t=document.getElementById(id); if(!t) return;
  var ths=t.tHead.querySelectorAll('tr.hdr th'); var head=[];
  for(var h=0;h<ths.length;h++){ head.push((ths[h].textContent||'').replace(/[\\u2191\\u2193\\u21C5]/g,'').trim()); }
  head.push('PeopleId'); head.push('OrgId');
  var lines=[head.map(paaCsvCell).join(',')];
  var rows=t.tBodies[0].rows;
  for(var r=0;r<rows.length;r++){ if(rows[r].style.display==='none') continue;
    var td=rows[r].cells, arr=[];
    for(var c=0;c<td.length;c++) arr.push((td[c].textContent||'').trim());
    arr.push(rows[r].getAttribute('data-pid')||''); arr.push(rows[r].getAttribute('data-oid')||'');
    lines.push(arr.map(paaCsvCell).join(','));
  }
  var blob=new Blob([lines.join('\\r\\n')],{type:'text/csv;charset=utf-8;'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=fname;
  document.body.appendChild(a); a.click(); setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(a.href);},100);
}
function paaTab(name){
  var panels=document.querySelectorAll('#paaBody > .tabpanel');
  for(var i=0;i<panels.length;i++){ panels[i].style.display=(name==='all'||panels[i].getAttribute('data-tab')===name)?'':'none'; }
  var tabs=document.querySelectorAll('.tabs .tab');
  for(var j=0;j<tabs.length;j++){ var isAll=tabs[j].getAttribute('data-tab')==='all';
    tabs[j].className='tab'+(isAll?' allt':'')+(tabs[j].getAttribute('data-tab')===name?' active':''); }
}
function paaPrint(){
  var node=document.querySelector('.paa'); if(!node){window.print();return;}
  var clone=node.cloneNode(true);
  var pn=clone.querySelectorAll('.tabpanel'); for(var p=0;p<pn.length;p++){ pn[p].style.display=''; }
  var kill=clone.querySelectorAll('.btn,.filters,.toolbar .chk,thead tr.flt,.tabs');
  for(var i=0;i<kill.length;i++){kill[i].parentNode.removeChild(kill[i]);}
  var rows=clone.querySelectorAll('tbody tr'); for(var r=0;r<rows.length;r++){ if(rows[r].style.display==='none') rows[r].parentNode.removeChild(rows[r]); }
  var css="*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}body{font-family:'Segoe UI',Arial,sans-serif;color:#1e293b;padding:20px}h2{color:#7f1d1d}.subject{background:#fef2f2;border-left:5px solid #b91c1c;padding:8px 12px;margin-bottom:10px}.subject .nm{font-size:16px;font-weight:800;color:#7f1d1d}.sec{font-size:14px;color:#7f1d1d;border-bottom:2px solid #7f1d1d;margin:12px 0 4px;font-weight:800}table{border-collapse:collapse;width:100%;font-size:11px}th{background:#eef2f7;text-align:left;padding:4px 8px}td{padding:4px 8px;border-bottom:1px solid #f1f5f9}.minorflag{color:#be123c;font-weight:800}.fam{color:#b45309;font-weight:700}.cards{display:flex;gap:10px;margin-bottom:12px}.card{border:1px solid #e2e8f0;border-top:3px solid #7f1d1d;border-radius:8px;padding:8px 12px}.card .val{font-size:20px;font-weight:800;color:#7f1d1d}.card .lbl{font-size:10px;text-transform:uppercase;color:#64748b}.arrow{display:none}a{color:#1e293b;text-decoration:none}";
  var pw=window.open('','_blank'); if(!pw){alert('Allow popups to print');return;}
  pw.document.write('<!DOCTYPE html><html><head><title>Person Attendance Audit</title><style>'+css+'</style></head><body>'+clone.innerHTML+'</body></html>');
  pw.document.close(); pw.focus(); setTimeout(function(){pw.print();},350);
}
</script>
"""


def window_form(pid):
    return ('<form method="get" class="filters">'
            '<input type="hidden" name="pid" value="%d">'
            '<label>Last N days<input type="number" name="days" min="1" placeholder="all history" value="%s"></label>'
            '<label>From<input type="date" name="from" value="%s"></label>'
            '<label>To<input type="date" name="to" value="%s"></label>'
            '<button type="submit" class="go">Apply window</button>'
            '<a class="reset" href="?pid=%d">reset</a>'
            '</form>' % (pid, (str(DAYS) if DAYS else ''), FROM, TO, pid))


def render_search():
    h = ['<form method="get" class="filters"><div class="search">'
         '<label style="min-width:280px;">Find a person<input type="text" name="q" value="%s" placeholder="last name, first name"></label>'
         '<button type="submit" class="go">Search</button></div></form>' % esc(QSEARCH)]
    if QSEARCH:
        term = QSEARCH.replace("'", "''")
        sql = ("SELECT TOP 50 p.PeopleId, p.Name2 AS Name, CONVERT(varchar(10),p.BDate,126) AS DOB, "
               "ISNULL(ms.Description,'') AS Status FROM dbo.People p "
               "LEFT JOIN lookup.MemberStatus ms ON ms.Id=p.MemberStatusId "
               "WHERE p.Name2 LIKE '%%%s%%' OR p.Name LIKE '%%%s%%' ORDER BY p.Name2" % (term, term))
        rows = list(q.QuerySql(sql))
        if not rows:
            h.append('<div class="muted">No matches for "%s".</div>' % esc(QSEARCH))
        else:
            h.append('<div class="sec">Matches <span class="c">(%d)</span></div><div class="rlist">' % len(rows))
            for r in rows:
                h.append('<a href="?pid=%d">%s <span class="muted">&middot; #%d &middot; DOB %s &middot; %s</span></a>'
                         % (r.PeopleId, esc(r.Name), r.PeopleId, esc(r.DOB or '?'), esc(r.Status)))
            h.append('</div>')
    else:
        h.append('<div class="muted">Enter a name to find the person to audit, or pass <code>?pid=</code> directly.</div>')
    return ''.join(h)


def status_label(deceased, archived, member_status):
    if deceased:
        return '<span class="minorflag">DECEASED</span>'
    if archived:
        return '<span class="fam">ARCHIVED</span>'
    return esc(member_status or '')


def subject_block(t):
    dob = t['DOB'] or '?'
    age = ('%d' % t['AgeNow']) if t['AgeNow'] is not None else '?'
    badge = ''
    if t.get('Deceased'):
        badge = ' <span class="minorflag">DECEASED</span>'
    elif t.get('Archived'):
        badge = ' <span class="fam">ARCHIVED</span>'
    return ('<div class="subject"><div class="nm">%s%s</div>'
            '<div class="meta">PeopleId %d &middot; %s &middot; DOB %s &middot; age %s now &middot; %s'
            ' &middot; <a href="/Person2/%d" target="_blank">open profile</a></div></div>'
            % (esc(t['Name']), badge, t['PeopleId'], esc(t['Gender'] or '?'), esc(dob), age,
               esc(t['MemberStatus'] or ''), t['PeopleId']))


def _now():
    try:
        return datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    except:
        return ''


def card(label, val, sub='', minor=False):
    cls = 'card minor' if minor else 'card'
    s = ('<div class="sub">%s</div>' % esc(sub)) if sub else ''
    return '<div class="%s"><div class="lbl">%s</div><div class="val">%d</div>%s</div>' % (cls, esc(label), val, s)


def kv_sheet(name, pairs):
    h = ['<table data-sheet="%s" style="display:none"><thead><tr class="hdr"><th>Field</th><th>Value</th></tr></thead><tbody>' % esc(name)]
    for k, v in pairs:
        h.append('<tr><td>%s</td><td>%s</td></tr>' % (esc(k), esc(v)))
    h.append('</tbody></table>')
    return ''.join(h)


def render_audit(pid):
    t = fetch_target(pid)
    if not t:
        return '<div class="cap">No person found for PeopleId %d.</div>' % pid

    sessions = fetch_sessions(pid)
    co = fetch_co(pid, t['FamilyId'])
    roles = fetch_roles(pid)
    vol, checks = fetch_background(pid)
    documents = fetch_documents(pid)
    access = fetch_access(pid)
    comms = fetch_comms(pid)
    checkouts = fetch_checkouts(pid)
    comms_personal_minor = sum(1 for x in comms if x['MinorRecipients'] >= 1 and x['Recipients'] <= 2)

    # per-meeting aggregates (for supervision ratio)
    others_by_meeting = {}
    minors_by_meeting = {}
    adults_by_meeting = {}
    for c in co:
        mid = c['MeetingId']
        others_by_meeting[mid] = others_by_meeting.get(mid, 0) + 1
        if c['Minor']:
            minors_by_meeting[mid] = minors_by_meeting.get(mid, 0) + 1
        elif c['Age'] is not None:
            adults_by_meeting[mid] = adults_by_meeting.get(mid, 0) + 1

    # per-person aggregates (span, sessions, age-at-attendance range)
    persons = {}
    for c in co:
        pp = persons.get(c['PeopleId'])
        if pp is None:
            pp = {'PeopleId': c['PeopleId'], 'Name': c['Name'], 'Gender': c['Gender'], 'DOB': c['DOB'],
                  'OrgId': c['OrgId'], 'SameFam': c['SameFam'], 'n': 0, 'dates': set(), 'ages': set(),
                  'orgs': set(), 'minor': False, 'Deceased': c['Deceased'], 'Archived': c['Archived'],
                  'MemberStatus': c['MemberStatus']}
            persons[c['PeopleId']] = pp
        pp['n'] += 1
        if c['Date']:
            pp['dates'].add(c['Date'])
        if c['Age'] is not None:
            pp['ages'].add(c['Age'])
        if c['Org']:
            pp['orgs'].add(c['Org'])
        if c['Minor']:
            pp['minor'] = True
    plist = list(persons.values())
    for pp in plist:
        ds = sorted(pp['dates'])
        pp['First'] = ds[0] if ds else ''
        pp['Last'] = ds[-1] if ds else ''
        pp['Span'] = days_between(pp['First'], pp['Last']) if ds else None

    distinct_people = len(persons)
    distinct_minors = len([pp for pp in plist if pp['minor']])
    minor_rows = sum(1 for c in co if c['Minor'])
    unknown_people = len(set(c['PeopleId'] for c in co if c['Age'] is None))
    solo_sessions = [s for s in sessions
                     if minors_by_meeting.get(s['MeetingId'], 0) >= 1 and adults_by_meeting.get(s['MeetingId'], 0) == 0]

    h = []
    h.append(subject_block(t))
    h.append(window_form(pid))
    h.append('<a class="muted" href="?">&larr; audit a different person</a>')

    h.append('<div class="toolbar" style="margin-top:8px;">'
             '<button class="btn" onclick="paaXls()">&#8681; Export Excel (all tabs)</button>'
             '<button class="btn alt" onclick="paaExport(\'paaCo\',\'coattendance.csv\')">&#8681; Co-attendance CSV</button>'
             '<button class="btn" onclick="paaPrint()">Print / Save</button></div>')

    h.append('<div class="cards">')
    h.append(card('Sessions present', len(sessions), WINDOW_LABEL))
    h.append(card('People co-present', distinct_people, '%d records' % len(co)))
    h.append(card('Distinct minors', distinct_minors, 'under %d at the time' % MINOR_AGE, minor=True))
    h.append(card('Solo with a minor', len(solo_sessions), 'no other adult present', minor=True))
    h.append(card('Minor co-attendances', minor_rows))
    h.append(card('Unknown age', unknown_people, 'no birthdate on file'))
    h.append('</div>')
    if len(co) >= CAP:
        h.append('<div class="cap">Row cap reached (%d). Narrow the date window for a complete extract.</div>' % CAP)

    # hidden Summary sheet (for Excel)
    h.append(kv_sheet('Summary', [
        ('Subject', t['Name']), ('PeopleId', t['PeopleId']), ('Gender', t['Gender']),
        ('DOB', t['DOB']), ('Age now', t['AgeNow']), ('Member status', t['MemberStatus']),
        ('Window', WINDOW_LABEL),
        ('Sessions present', len(sessions)), ('People co-present', distinct_people),
        ('Co-attendance records', len(co)), ('Distinct minors', distinct_minors),
        ('Minor co-attendances', minor_rows), ('Solo-with-minor sessions', len(solo_sessions)),
        ('Unknown-age people', unknown_people),
        ('Volunteer clearance', (vol['Status'] if vol else 'no record')),
        ('Children-cleared', ('yes' if vol and vol['Children'] else 'no')),
        ('Leader-cleared', ('yes' if vol and vol['Leader'] else 'no')),
        ('Background checks on file', len(checks)),
        ('Latest background check', (checks[0]['Created'] if checks else '')),
        ('Documents on file', len(documents)),
        ('TouchPoint login', (access[0]['Username'] if access else 'none')),
        ('TouchPoint roles', ('; '.join(a['Roles'] or '' for a in access) if access else 'none')),
        ('In-system messages', len(comms)),
        ('Messages reaching a minor (1-2 recipients)', comms_personal_minor),
        ('Deceased', ('yes' if t.get('Deceased') else 'no')),
        ('Archived', ('yes' if t.get('Archived') else 'no')),
        ('Generated', _now()),
    ]))

    # on-screen tab navigation (Excel/CSV/Print always include everything)
    h.append('<div class="tabs">'
             '<a class="tab active" data-tab="background" onclick="paaTab(\'background\');return false;" href="#">Background</a>'
             '<a class="tab" data-tab="people" onclick="paaTab(\'people\');return false;" href="#">People</a>'
             '<a class="tab" data-tab="sessions" onclick="paaTab(\'sessions\');return false;" href="#">Sessions</a>'
             '<a class="tab" data-tab="checkouts" onclick="paaTab(\'checkouts\');return false;" href="#">Check-outs</a>'
             '<a class="tab" data-tab="comms" onclick="paaTab(\'comms\');return false;" href="#">Comms</a>'
             '<a class="tab" data-tab="coattend" onclick="paaTab(\'coattend\');return false;" href="#">Co-attendance</a>'
             '<a class="tab allt" data-tab="all" onclick="paaTab(\'all\');return false;" href="#">All (print)</a>'
             '</div><div id="paaBody">')

    # ---- roles & background ----
    h.append('<div class="tabpanel" data-tab="background">')
    h.append('<div class="sec">Roles &amp; background</div>')
    if vol:
        flags = []
        if vol['Standard']:
            flags.append('Standard')
        if vol['Children']:
            flags.append('Children')
        if vol['Leader']:
            flags.append('Leader')
        h.append('<div class="subject" style="background:#f0f9ff;border-color:#bae6fd;border-left-color:#0369a1;">'
                 '<div class="meta" style="color:#0c4a6e;font-size:13px;"><b>Volunteer clearance:</b> %s%s'
                 ' &middot; processed %s &middot; MVR %s &middot; training %s</div></div>'
                 % (esc(vol['Status']), (' &middot; cleared for: ' + esc(', '.join(flags))) if flags else '',
                    esc(vol['Processed'] or '?'), esc(vol['MVR'] or '?'), esc(vol['Trained'] or '?')))
    else:
        h.append('<div class="muted">No volunteer/clearance record on file.</div>')
    if checks:
        h.append('<table data-sheet="Background Checks"><thead><tr class="hdr">'
                 '<th>Created</th><th>Approval</th><th>Level</th><th>Child-serving</th></tr></thead><tbody>')
        for ck in checks:
            h.append('<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                     % (ck['Created'], esc(ck['Approval']), esc(ck['Level']), ('yes' if ck['ChildServing'] else '')))
        h.append('</tbody></table>')
    if roles:
        h.append('<table data-sheet="Roles" id="paaRoles"><thead><tr class="hdr">')
        for i, cn in enumerate(['Organization', 'Role', 'Org status', 'Enrolled', 'Inactive']):
            h.append('<th onclick="paaSort(\'paaRoles\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
        h.append('</tr></thead><tbody>')
        for rr in roles:
            rolecell = ('<b>' + esc(rr['Role']) + '</b>') if rr['IsLeader'] else esc(rr['Role'])
            h.append('<tr data-oid="%d"><td><a href="/Organization/%d" target="_blank">%s</a></td>'
                     '<td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                     % (rr['OrgId'], rr['OrgId'], esc(rr['Org']), rolecell, esc(rr['OrgStatus']),
                        rr['Enrolled'], rr['Inactive']))
        h.append('</tbody></table>')
    else:
        h.append('<div class="muted">No involvement memberships on file.</div>')

    # ---- documents on file (all, no name filter) ----
    h.append('<div class="sec">Documents on file <span class="c">(%d)</span></div>' % len(documents))
    if not documents:
        h.append('<div class="muted">No documents / forms on file.</div>')
    else:
        h.append('<table id="paaDocs" data-sheet="Documents"><thead><tr class="hdr">')
        for i, cn in enumerate(['Date', 'Document', 'Type', 'Uploaded by']):
            h.append('<th onclick="paaSort(\'paaDocs\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
        h.append('</tr></thead><tbody>')
        for dcu in documents:
            typ = 'document' if dcu['IsDoc'] else 'form'
            namecell = esc(dcu['Name'])
            if dcu['LargeId']:
                namecell = '<a href="/Display/%d" target="_blank">%s</a>' % (dcu['LargeId'], namecell)
            h.append('<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                     % (dcu['AppDate'], namecell, typ, esc(dcu['Uploader'])))
        h.append('</tbody></table>')

    # ---- touchpoint access (login + roles) ----
    h.append('<div class="sec">TouchPoint access <span class="c">(login &amp; permissions)</span></div>')
    if not access:
        h.append('<div class="muted">No TouchPoint login on file for this person.</div>')
    else:
        h.append('<table data-sheet="TouchPoint Access"><thead><tr class="hdr">'
                 '<th>Username</th><th>Roles (permissions)</th><th>Last login</th><th>Created</th><th>Locked</th></tr></thead><tbody>')
        for u in access:
            h.append('<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                     % (esc(u['Username']), esc(u['Roles'] or '(none)'), u['LastLogin'] or 'never',
                        u['Created'], ('<span class="minorflag">LOCKED</span>' if u['Locked'] else '')))
        h.append('</tbody></table>')

    # ---- people co-present: summary (span + age-at-attendance range) ----
    h.append('</div><div class="tabpanel" data-tab="people">')
    h.append('<div class="sec">People co-present &mdash; summary '
             '<span class="c">(%d people &middot; how long &amp; how often)</span></div>' % distinct_people)
    h.append('<div class="toolbar">'
             '<button class="btn alt" onclick="paaExport(\'paaPeople\',\'people_summary.csv\')">&#8681; Export CSV</button>'
             '<label class="chk"><input type="checkbox" id="hide_paaPeople" onchange="paaFilter(\'paaPeople\')"> Minors only</label>'
             '<span class="count"><b id="shown_paaPeople">%d</b> shown</span></div>' % distinct_people)
    h.append('<table id="paaPeople" data-sheet="People Summary"><thead><tr class="hdr">')
    for i, cn in enumerate(['Person', 'Minor', 'Sessions', 'First', 'Last', 'Span', 'Age at attendance', 'Gender', 'Status', 'Where']):
        h.append('<th onclick="paaSort(\'paaPeople\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
    h.append('</tr><tr class="flt">')
    for i in range(10):
        h.append('<th><input type="text" oninput="paaFilter(\'paaPeople\')" placeholder="filter"></th>')
    h.append('</tr></thead><tbody>')
    for pp in sorted(plist, key=lambda x: (-x['n'], x['Name'])):
        ages = sorted(pp['ages'])
        arange = ('%d-%d' % (ages[0], ages[-1])) if len(ages) > 1 else ('%d' % ages[0]) if ages else '?'
        fam = ' <span class="fam">FAM</span>' if pp['SameFam'] else ''
        minorcell = '<span class="minorflag">MINOR</span>' if pp['minor'] else ''
        orgs = ', '.join(sorted(pp['orgs']))
        h.append('<tr data-pid="%d" data-oid="%d" data-minor="%d">'
                 '<td><a href="?pid=%d&amp;co=%d">%s</a>%s'
                 ' <a href="/Person2/%d" target="_blank" class="muted" style="font-size:11px;">&#9432;</a></td>'
                 '<td>%s</td><td class="num">%d</td><td>%s</td><td>%s</td><td>%s</td>'
                 '<td class="num">%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                 % (pp['PeopleId'], pp['OrgId'], (1 if pp['minor'] else 0),
                    pid, pp['PeopleId'], esc(pp['Name']), fam, pp['PeopleId'],
                    minorcell, pp['n'], pp['First'], pp['Last'], dur_label(pp['Span']),
                    arange, esc(pp['Gender'] or ''),
                    status_label(pp['Deceased'], pp['Archived'], pp['MemberStatus']), esc(orgs)))
    h.append('</tbody></table>')

    # ---- sessions present (supervision ratio) ----
    h.append('</div><div class="tabpanel" data-tab="sessions">')
    h.append('<div class="sec">Sessions present <span class="c">(%d meetings &middot; %d with the subject alone with a minor)</span></div>'
             % (len(sessions), len(solo_sessions)))
    if sessions:
        sess_sorted = sorted(sessions, key=lambda s: s['Date'], reverse=True)
        h.append('<table id="paaSess" data-sheet="Sessions"><thead><tr class="hdr">')
        for i, cn in enumerate(['Date', 'Day', 'Organization', 'Target role', 'Adults', 'Minors', 'Flag']):
            h.append('<th onclick="paaSort(\'paaSess\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
        h.append('</tr></thead><tbody>')
        for s in sess_sorted:
            mid = s['MeetingId']
            mn = minors_by_meeting.get(mid, 0)
            ad = adults_by_meeting.get(mid, 0)
            flag = '<span class="minorflag">SOLO w/ MINOR</span>' if (mn >= 1 and ad == 0) else (
                'low adult' if (mn >= 1 and ad == 1) else '')
            h.append('<tr data-oid="%d"><td>%s</td><td>%s</td>'
                     '<td><a href="/Organization/%d" target="_blank">%s</a></td>'
                     '<td>%s</td><td class="num">%d</td><td class="num %s">%d</td><td>%s</td></tr>'
                     % (s['OrgId'], s['Date'], weekday(s['Date']), s['OrgId'], esc(s['Org']),
                        esc(s['Role'] or ''), ad, 'minorflag' if mn else '', mn, flag))
        h.append('</tbody></table>')

    # ---- check-outs by the subject ----
    h.append('</div><div class="tabpanel" data-tab="checkouts">')
    h.append('<div class="sec">Check-outs by subject '
             '<span class="c">(%d records &middot; %d involving a minor)</span></div>'
             % (len(checkouts), sum(1 for x in checkouts if x['Minor'])))
    if not checkouts:
        h.append('<div class="muted">No records where this person checked/picked someone out.</div>')
    else:
        h.append('<table id="paaCheckouts" data-sheet="Check-outs"><thead><tr class="hdr">')
        for i, cn in enumerate(['Date', 'Day', 'Person', 'Age', 'Minor', 'Organization']):
            h.append('<th onclick="paaSort(\'paaCheckouts\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
        h.append('</tr></thead><tbody>')
        for x in checkouts:
            agecell = ('%d' % x['Age']) if x['Age'] is not None else '?'
            minorcell = '<span class="minorflag">MINOR</span>' if x['Minor'] else ''
            h.append('<tr data-pid="%d" data-oid="0"><td>%s</td><td>%s</td>'
                     '<td><a href="/Person2/%d" target="_blank">%s</a></td>'
                     '<td class="num">%s</td><td class="num">%s</td><td>%s</td></tr>'
                     % (x['PeopleId'], x['Date'], weekday(x['Date']), x['PeopleId'], esc(x['Name']),
                        agecell, minorcell, esc(x['Org'])))
        h.append('</tbody></table>')

    # ---- in-system communications (email/text; platform only) ----
    h.append('</div><div class="tabpanel" data-tab="comms">')
    h.append('<div class="sec">In-system communications '
             '<span class="c">(%d &middot; email + text; %d reached a minor with 1&ndash;2 recipients)</span></div>'
             % (len(comms), comms_personal_minor))
    h.append('<div class="muted" style="margin-bottom:4px;">Platform messaging only and <b>mostly bulk</b>. '
             'A 1:1 (or 1&ndash;2 recipient) message that reached a minor is flagged &mdash; treat as a lead, not proof. '
             'Personal off-platform contact (cell, personal email, DMs) is NOT captured.</div>')
    if not comms:
        h.append('<div class="muted">No in-system messages on file.</div>')
    else:
        h.append('<div class="toolbar">'
                 '<button class="btn alt" onclick="paaExport(\'paaComms\',\'communications.csv\')">&#8681; Export CSV</button>'
                 '<label class="chk"><input type="checkbox" id="hide_paaComms" onchange="paaFilter(\'paaComms\')"> Reached a minor (1&ndash;2 recip) only</label>'
                 '<span class="count"><b id="shown_paaComms">%d</b> shown</span></div>' % len(comms))
        h.append('<table id="paaComms" data-sheet="Communications"><thead><tr class="hdr">')
        for i, cn in enumerate(['Date', 'Channel', 'Recipients', 'Minor recip', 'Type', 'Summary']):
            h.append('<th onclick="paaSort(\'paaComms\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
        h.append('</tr><tr class="flt">')
        for i in range(6):
            h.append('<th><input type="text" oninput="paaFilter(\'paaComms\')" placeholder="filter"></th>')
        h.append('</tr></thead><tbody>')
        for x in comms:
            personal = (x['MinorRecipients'] >= 1 and x['Recipients'] <= 2)
            flag = ' <span class="minorflag">MINOR</span>' if personal else ''
            h.append('<tr data-minor="%d"><td>%s</td><td>%s</td><td class="num">%d</td>'
                     '<td class="num %s">%d</td><td>%s</td><td>%s%s</td></tr>'
                     % ((1 if personal else 0), x['Date'], esc(x['Channel']), x['Recipients'],
                        'minorflag' if x['MinorRecipients'] else '', x['MinorRecipients'],
                        esc(x['Type']), esc(x['Summary']), flag))
        h.append('</tbody></table>')

    # ---- full co-attendance detail (per session, raw evidence) ----
    h.append('</div><div class="tabpanel" data-tab="coattend">')
    h.append('<div class="sec">Co-attendance detail <span class="c">(%d records, %d people)</span></div>'
             % (len(co), distinct_people))
    h.append('<div class="toolbar">'
             '<button class="btn alt" onclick="paaExport(\'paaCo\',\'coattendance.csv\')">&#8681; Export CSV</button>'
             '<label class="chk"><input type="checkbox" id="hide_paaCo" onchange="paaFilter(\'paaCo\')"> Minors only</label>'
             '<span class="count"><b id="shown_paaCo">%d</b> shown</span></div>' % len(co))
    h.append('<table id="paaCo" data-sheet="Co-Attendance"><thead><tr class="hdr">')
    for i, cn in enumerate(['Date', 'Day', 'Organization', 'Person', 'Age', 'Minor', 'Role', 'Gender', 'DOB', 'Status']):
        h.append('<th onclick="paaSort(\'paaCo\',%d)">%s <span class="arrow">&#8645;</span></th>' % (i, cn))
    h.append('</tr><tr class="flt">')
    for i in range(10):
        h.append('<th><input type="text" oninput="paaFilter(\'paaCo\')" placeholder="filter"></th>')
    h.append('</tr></thead><tbody>')
    for c in co:
        minorcell = '<span class="minorflag">MINOR</span>' if c['Minor'] else ''
        fam = ' <span class="fam">FAM</span>' if c['SameFam'] else ''
        agecell = ('%d' % c['Age']) if c['Age'] is not None else '?'
        h.append('<tr data-pid="%d" data-oid="%d" data-minor="%d">'
                 '<td>%s</td><td>%s</td>'
                 '<td><a href="/Organization/%d" target="_blank">%s</a></td>'
                 '<td><a href="?pid=%d&amp;co=%d">%s</a>%s'
                 ' <a href="/Person2/%d" target="_blank" class="muted" style="font-size:11px;">&#9432;</a></td>'
                 '<td class="num">%s</td><td class="num">%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                 % (c['PeopleId'], c['OrgId'], (1 if c['Minor'] else 0),
                    c['Date'], weekday(c['Date']), c['OrgId'], esc(c['Org']),
                    pid, c['PeopleId'], esc(c['Name']), fam, c['PeopleId'],
                    agecell, minorcell, esc(c['Role'] or ''), esc(c['Gender'] or ''), esc(c['DOB'] or ''),
                    status_label(c['Deceased'], c['Archived'], c['MemberStatus'])))
    h.append('</tbody></table>')
    h.append('</div>')   # close coattend panel
    h.append('</div>')   # close #paaBody

    # hidden Method & Caveats sheet
    h.append(kv_sheet('Method', [
        ('Source', 'TouchPoint Attend + Meetings (AttendanceFlag=1, DidNotMeet=0)'),
        ('Co-attendance', 'Everyone else present at a meeting the subject was present for'),
        ('Age at attendance', 'Age on the meeting date'),
        ('Minor', 'Under %d at the time of attendance' % MINOR_AGE),
        ('Solo-with-minor', 'A session with 1+ minor and no other known adult present'),
        ('Span', 'First to last co-attended date (not thresholded)'),
        ('Check-outs', 'Attend rows where CheckOutBy = the subject (they picked the person up)'),
        ('In-system comms', 'Email/text the subject SENT (EmailQueue.QueuedBy / SMSList.SenderID) + inbound '
                            'texts (SmsReceived.FromPeopleId). Recipient counts separate 1:1 from bulk.'),
        ('CAVEAT - attendance', 'Only marked attendance is counted. Unmarked meetings, off-campus / informal '
                                'contact, and people without a birthdate (unknown age, not minor-flagged) are NOT captured.'),
        ('CAVEAT - comms', 'Platform messaging ONLY and mostly bulk broadcasts. Personal off-platform contact '
                           '(personal cell, personal email, DMs, social media) is NOT captured. Treat as a lead, not proof.'),
        ('Generated', _now()),
        ('Scope', 'Read-only. Authorized safe-church / incident review only.'),
    ]))
    h.append('<script>paaTab("background");</script>')
    return ''.join(h)


def render_drill(pid, co_pid):
    t = fetch_target(pid)
    if not t:
        return '<div class="cap">No person found.</div>'
    co = [c for c in fetch_co(pid, t['FamilyId']) if c['PeopleId'] == co_pid]
    name = co[0]['Name'] if co else ('Person %d' % co_pid)
    h = [subject_block(t)]
    h.append('<a class="muted" href="?pid=%d">&larr; back to audit</a>' % pid)
    h.append('<div class="sec">Sessions <a href="/Person2/%d" target="_blank">%s</a> was co-present with the subject '
             '<span class="c">(%d)</span></div>' % (co_pid, esc(name), len(co)))
    if not co:
        h.append('<div class="muted">No co-attended sessions found in this window.</div>')
        return ''.join(h)
    h.append('<button class="btn" onclick="paaPrint()">Print / Save</button>')
    h.append('<table><thead><tr><th>Date</th><th>Day</th><th>Organization</th><th class="num">Their age</th><th>Their role</th></tr></thead><tbody>')
    for c in sorted(co, key=lambda x: x['Date']):
        agecell = ('%d' % c['Age']) if c['Age'] is not None else '?'
        cls = ' class="num minorflag"' if c['Minor'] else ' class="num"'
        h.append('<tr><td>%s</td><td>%s</td><td><a href="/Organization/%d" target="_blank">%s</a></td>'
                 '<td%s>%s</td><td>%s</td></tr>'
                 % (c['Date'], weekday(c['Date']), c['OrgId'], esc(c['Org']), cls, agecell, esc(c['Role'] or '')))
    h.append('</tbody></table>')
    return ''.join(h)


print CSS
print PRINT_JS
print '<div class="paa">'
print '<h2>Person Attendance Audit</h2>'
print ('<p class="lede">Point this at one person to reconstruct, from actual attendance, every session they were '
       'present for and <b>everyone else present at those same sessions</b> &mdash; with each person&#39;s age at '
       'the time and a minor flag. Authorized safe-church / incident review only. Read-only.</p>')
try:
    if PID and CO:
        print render_drill(PID, CO)
    elif PID:
        print render_audit(PID)
    else:
        print render_search()
except Exception as e:
    print '<div style="color:#b91c1c;font-family:monospace;white-space:pre-wrap;">Error: {0}\n{1}</div>'.format(str(e), traceback.format_exc())
print '</div>'
