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



# ============================================================
# TPxi_ProgramPulse.py
# Program Pulse - Program Activity Summary for Leaders
# ============================================================
# Generates scheduled digest reports summarizing activity
# across selected programs: new enrollments, drops, baptisms,
# attendance gaps, transactions, and more.
#
# Dashboard: /PyScriptForm/TPxi_ProgramPulse
# Batch:     /PyScript/TPxi_ProgramPulse_Batch
# ============================================================
#
# INSTALLATION
# ============================================================
#
# Step 1: Add this script
#   - Admin > Advanced > Special Content > Python
#   - Click "Add New", name it: TPxi_ProgramPulse
#   - Paste this entire file and Save
#
# Step 2: Access the dashboard
#   - Navigate to /PyScript/TPxi_ProgramPulse
#   - Set a default sender in Settings
#   - Create your first report configuration
#
# Step 3: Enable automatic sending (optional)
#   - Open your MorningBatch script in Special Content > Python
#   - Add these two lines:
#
#       Data.run_batch = "true"
#       model.CallScript("TPxi_ProgramPulse")
#
#   - Reports will only send on each config's scheduled day
#
# CONTENT STORAGE
# ============================================================
# Program Pulse stores data in these auto-created text entries:
#   ProgramPulse_Configs   - Report configurations
#   ProgramPulse_Settings  - Global settings (sender, contact methods)
#   ProgramPulse_Log       - Send history (last 100 entries)
# These persist independently of the script code.
# ============================================================

import json
import datetime
import traceback

# ============================================================
# CONFIGURATION
# ============================================================
CONTENT_KEY = "ProgramPulse_Configs"
SETTINGS_KEY = "ProgramPulse_Settings"
LOG_KEY = "ProgramPulse_Log"
TITLE = "Program Pulse"
CSS_PREFIX = "pp"
DEFAULT_LOOKBACK_DAYS = 7
STALE_THRESHOLD_DAYS = 30
ATTENDANCE_GAP_WEEKS = 4

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def get_form_data(attr_name, default_value):
    try:
        value = getattr(model.Data, attr_name, None)
        if value is None or str(value).strip() == '':
            return default_value
        return str(value).strip()
    except:
        return default_value

def load_configs():
    try:
        content = model.TextContent(CONTENT_KEY)
        if content:
            return json.loads(content)
    except:
        pass
    return []

def save_configs(configs):
    model.WriteContentText(CONTENT_KEY, json.dumps(configs, indent=2), "")

def load_settings():
    try:
        content = model.TextContent(SETTINGS_KEY)
        if content:
            return json.loads(content)
    except:
        pass
    return {}

def save_settings(settings):
    model.WriteContentText(SETTINGS_KEY, json.dumps(settings, indent=2), "")

def get_contact_methods_config():
    """Get contact methods from PP settings."""
    settings = load_settings()
    return settings.get('contact_methods', [])

def load_log():
    try:
        content = model.TextContent(LOG_KEY)
        if content:
            return json.loads(content)
    except:
        pass
    return []

def append_log(entry):
    logs = load_log()
    logs.append(entry)
    logs = logs[-100:]
    try:
        model.WriteContentText(LOG_KEY, json.dumps(logs, indent=2), "")
    except:
        pass

def generate_id():
    now = datetime.datetime.now()
    return "pp_{0}".format(now.strftime("%Y%m%d%H%M%S"))

def safe_int(val, default=0):
    try:
        return int(val)
    except:
        return default

def format_date(dt_str):
    if not dt_str:
        return ""
    try:
        if hasattr(dt_str, 'strftime'):
            return dt_str.strftime("%m/%d/%Y")
        s = str(dt_str)
        if 'T' in s:
            s = s.split('T')[0]
        parts = s.split('-')
        if len(parts) == 3:
            return "{0}/{1}/{2}".format(parts[1], parts[2], parts[0])
        return s
    except:
        return str(dt_str)

def safe_str(val):
    if val is None:
        return ""
    try:
        s = str(val)
        result = []
        for c in s:
            try:
                c.encode('ascii')
                result.append(c)
            except:
                result.append('?')
        return ''.join(result)
    except:
        return ""

def get_prog_ids_from_config(config):
    prog_div_groups = config.get('programDivGroups', [])
    prog_ids = set()
    for g in prog_div_groups:
        pid = g.get('programId')
        if pid:
            prog_ids.add(int(pid))
    return list(prog_ids)

def get_org_ids_from_config(config):
    prog_div_groups = config.get('programDivGroups', [])
    exclude_ids = config.get('excludeOrgIds', '')
    exclude_set = set()
    if exclude_ids:
        for oid in str(exclude_ids).split(','):
            oid = oid.strip()
            if oid.isdigit():
                exclude_set.add(int(oid))

    conditions = []
    for g in prog_div_groups:
        pid = g.get('programId')
        did = g.get('divisionId')
        if pid and did:
            conditions.append("(d.ProgId = {0} AND o.DivisionId = {1})".format(pid, did))
        elif pid:
            conditions.append("(d.ProgId = {0})".format(pid))

    if not conditions:
        return []

    where = " OR ".join(conditions)
    sql = """
        SELECT o.OrganizationId
        FROM Organizations o
        JOIN Division d ON o.DivisionId = d.Id
        WHERE ({0})
        AND o.OrganizationStatusId = 30
    """.format(where)

    org_ids = []
    try:
        for row in q.QuerySql(sql):
            oid = row.OrganizationId
            if oid not in exclude_set:
                org_ids.append(oid)
    except:
        pass
    return org_ids


# ============================================================
# REPORT GENERATOR
# ============================================================

class ReportGenerator:
    def __init__(self, config, start_date, end_date):
        self.config = config
        self.start_date = start_date
        self.end_date = end_date
        self.org_ids = get_org_ids_from_config(config)
        self.sections_config = config.get('sections', {})

    def _section_level(self, key):
        """Get section level: 'none', 'summary', or 'detail'. Backwards compatible with bool."""
        val = self.sections_config.get(key, 'detail')
        if val is True:
            return 'detail'
        if val is False or val == 'none':
            return 'none'
        return str(val)

    def generate(self):
        if not self.org_ids:
            return "<p>No organizations found for the selected programs/divisions.</p>"

        sections = []
        section_map = [
            ('newEnrollments', 'New Enrollments', self.new_enrollments),
            ('droppedMembers', 'Dropped Members', self.dropped_members),
            ('newInvolvements', 'New Involvements Created', self.new_involvements),
            ('staleInvolvements', 'Stale Involvements', self.stale_involvements),
            ('transactions', 'Transactions', self.transactions),
            ('newPeople', 'New People', self.new_people),
            ('newBaptisms', 'New Baptisms', self.new_baptisms),
            ('newChurchMembers', 'New Church Members', self.new_church_members),
            ('attendanceGaps', 'Attendance Gaps', self.attendance_gaps),
        ]

        hide_empty = self.config.get('hideEmptySections', False)

        for key, title, func in section_map:
            level = self._section_level(key)
            if level != 'none':
                try:
                    content = func(level)
                    if hide_empty and content.startswith(self.EMPTY_SENTINEL):
                        continue
                    sections.append((title, content))
                except Exception as e:
                    sections.append((title, "<p style='color:#c00;'>Error: {0}</p>".format(safe_str(e))))

        return self.build_report(sections)

    def _org_id_list(self):
        return ','.join(str(x) for x in self.org_ids)

    def new_involvements(self, level='detail'):
        sql = """
            SELECT o.OrganizationId, o.OrganizationName, o.CreatedDate,
                   p.Name as Program, d.Name as Division,
                   ISNULL(u.Name, '') as CreatorName
            FROM Organizations o
            LEFT JOIN Division d ON o.DivisionId = d.Id
            LEFT JOIN Program p ON d.ProgId = p.Id
            LEFT JOIN Users u ON CASE WHEN ISNUMERIC(CAST(o.CreatedBy AS VARCHAR(50))) = 1 THEN CAST(o.CreatedBy AS INT) END = u.UserId
            WHERE o.OrganizationId IN ({0})
            AND o.CreatedDate >= '{1}' AND o.CreatedDate <= '{2}'
            ORDER BY o.CreatedDate DESC
        """.format(self._org_id_list(), self.start_date, self.end_date)
        rows = list(q.QuerySql(sql))
        if not rows:
            return self._empty("No new involvements created this period.")
        if level == 'summary':
            return self._summary_count(len(rows), "new involvement", "new involvements")
        headers = ['Organization', 'Program', 'Division', 'Created', 'Created By']
        data = []
        for r in rows:
            data.append([
                '<a href="/Organization/{0}" target="_blank">{1}</a>'.format(r.OrganizationId, safe_str(r.OrganizationName)),
                safe_str(r.Program),
                safe_str(r.Division),
                format_date(r.CreatedDate),
                safe_str(r.CreatorName) if r.CreatorName else ''
            ])
        return self._table(headers, data, len(rows))

    def stale_involvements(self, level='detail'):
        threshold = self.config.get('staleThresholdDays', STALE_THRESHOLD_DAYS)
        sql = """
            SELECT o.OrganizationId, o.OrganizationName,
                   p.Name as Program, d.Name as Division,
                   o.CreatedDate, ISNULL(u.Name, '') as CreatorName,
                   MAX(m.MeetingDate) as LastMeeting,
                   DATEDIFF(day, MAX(m.MeetingDate), GETDATE()) as DaysSince
            FROM Organizations o
            LEFT JOIN Division d ON o.DivisionId = d.Id
            LEFT JOIN Program p ON d.ProgId = p.Id
            LEFT JOIN Meetings m ON o.OrganizationId = m.OrganizationId AND m.DidNotMeet = 0
            LEFT JOIN Users u ON CASE WHEN ISNUMERIC(CAST(o.CreatedBy AS VARCHAR(50))) = 1 THEN CAST(o.CreatedBy AS INT) END = u.UserId
            WHERE o.OrganizationId IN ({0})
            AND o.OrganizationStatusId = 30
            GROUP BY o.OrganizationId, o.OrganizationName, p.Name, d.Name,
                     o.CreatedDate, u.Name
            HAVING MAX(m.MeetingDate) IS NULL
                OR DATEDIFF(day, MAX(m.MeetingDate), GETDATE()) > {1}
            ORDER BY p.Name, d.Name, o.OrganizationName
        """.format(self._org_id_list(), threshold)
        rows = list(q.QuerySql(sql))
        if not rows:
            return self._empty("No stale involvements found.")
        if level == 'summary':
            return self._summary_count(len(rows), "stale involvement", "stale involvements")

        # Group by Program > Division
        groups = {}
        group_order = []
        for r in rows:
            prog = safe_str(r.Program) if r.Program else 'No Program'
            div = safe_str(r.Division) if r.Division else 'No Division'
            key = prog + ' > ' + div
            if key not in groups:
                groups[key] = {'program': prog, 'division': div, 'orgs': []}
                group_order.append(key)
            groups[key]['orgs'].append(r)

        html = '<div style="margin-bottom:8px;color:#666;font-size:13px;">{0} stale involvement{1}</div>'.format(len(rows), 's' if len(rows) != 1 else '')

        for key in group_order:
            grp = groups[key]
            html += '<div style="margin-bottom:16px;border:1px solid #eee;border-radius:6px;overflow:hidden;">'
            html += '<div style="padding:10px 14px;background:#f8f9fa;border-bottom:1px solid #eee;">'
            html += '<span style="font-weight:600;font-size:14px;">{0}</span>'.format(grp['program'])
            html += ' <span style="color:#888;font-size:13px;">&rsaquo; {0}</span>'.format(grp['division'])
            html += ' <span style="font-size:12px;color:#666;margin-left:8px;">({0} involvement{1})</span>'.format(len(grp['orgs']), 's' if len(grp['orgs']) != 1 else '')
            html += '</div>'

            html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
            html += '<thead><tr>'
            html += '<th style="text-align:left;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">INVOLVEMENT</th>'
            html += '<th style="text-align:left;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">CREATED</th>'
            html += '<th style="text-align:left;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">CREATED BY</th>'
            html += '<th style="text-align:right;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">LAST MEETING</th>'
            html += '<th style="text-align:right;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">DAYS</th>'
            html += '</tr></thead><tbody>'
            for i, r in enumerate(grp['orgs']):
                bg = '#fff' if i % 2 == 0 else '#f8f9fa'
                last = format_date(r.LastMeeting) if r.LastMeeting else 'Never'
                days = r.DaysSince if r.DaysSince else 'N/A'
                days_color = '#e74c3c' if r.DaysSince and r.DaysSince > 90 else '#f39c12' if r.DaysSince and r.DaysSince > 60 else '#666'
                html += '<tr style="background:{0};">'.format(bg)
                created = ''
                if r.CreatedDate:
                    try:
                        created = r.CreatedDate.strftime("%m/%d/%Y") if hasattr(r.CreatedDate, 'strftime') else str(r.CreatedDate).split(' ')[0].split('T')[0]
                    except:
                        created = format_date(r.CreatedDate)
                html += '<td style="padding:5px 14px;"><a href="/Organization/{0}" target="_blank">{1}</a></td>'.format(r.OrganizationId, safe_str(r.OrganizationName))
                html += '<td style="padding:5px 14px;">{0}</td>'.format(created)
                html += '<td style="padding:5px 14px;">{0}</td>'.format(safe_str(r.CreatorName) if r.CreatorName else '')
                html += '<td style="padding:5px 14px;text-align:right;">{0}</td>'.format(last)
                html += '<td style="padding:5px 14px;text-align:right;font-weight:600;color:{0};">{1}</td>'.format(days_color, str(days))
                html += '</tr>'
            html += '</tbody></table></div>'

        return html

    def new_enrollments(self, level='detail'):
        sql = """
            SELECT TOP 500 p.PeopleId, p.Name2, o.OrganizationId, o.OrganizationName,
                   om.EnrollmentDate, mt.Description as MemberType,
                   prog.Name as Program, d.Name as Division
            FROM OrganizationMembers om
            JOIN People p ON om.PeopleId = p.PeopleId
            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
            LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
            LEFT JOIN Division d ON o.DivisionId = d.Id
            LEFT JOIN Program prog ON d.ProgId = prog.Id
            WHERE o.OrganizationId IN ({0})
            AND om.EnrollmentDate >= '{1}' AND om.EnrollmentDate <= '{2}'
            AND p.IsDeceased = 0 AND ISNULL(p.ArchivedFlag, 0) = 0
            ORDER BY p.Name2
        """.format(self._org_id_list(), self.start_date, self.end_date)
        rows = list(q.QuerySql(sql))
        if not rows:
            return self._empty("No new enrollments this period.")
        if level == 'summary':
            return self._org_summary(rows, len(rows), "new enrollment", "new enrollments")
        headers = ['Person', 'Organization', 'Division', 'Member Type', 'Enrolled']
        data = []
        for r in rows:
            data.append([
                '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(r.PeopleId, safe_str(r.Name2)),
                '<a href="/Organization/{0}" target="_blank">{1}</a>'.format(r.OrganizationId, safe_str(r.OrganizationName)),
                safe_str(r.Division),
                safe_str(r.MemberType),
                format_date(r.EnrollmentDate)
            ])
        return self._table(headers, data, len(rows))

    def dropped_members(self, level='detail'):
        sql = """
            SELECT TOP 500 p.PeopleId, p.Name2, o.OrganizationId, o.OrganizationName,
                   om.InactiveDate, mt.Description as MemberType,
                   prog.Name as Program, d.Name as Division
            FROM OrganizationMembers om
            JOIN People p ON om.PeopleId = p.PeopleId
            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
            LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
            LEFT JOIN Division d ON o.DivisionId = d.Id
            LEFT JOIN Program prog ON d.ProgId = prog.Id
            WHERE o.OrganizationId IN ({0})
            AND om.InactiveDate >= '{1}' AND om.InactiveDate <= '{2}'
            AND p.IsDeceased = 0 AND ISNULL(p.ArchivedFlag, 0) = 0
            ORDER BY om.InactiveDate DESC
        """.format(self._org_id_list(), self.start_date, self.end_date)
        rows = list(q.QuerySql(sql))
        if not rows:
            return self._empty("No dropped members this period.")
        if level == 'summary':
            return self._org_summary(rows, len(rows), "dropped member", "dropped members")
        headers = ['Person', 'Organization', 'Division', 'Member Type', 'Dropped']
        data = []
        for r in rows:
            data.append([
                '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(r.PeopleId, safe_str(r.Name2)),
                '<a href="/Organization/{0}" target="_blank">{1}</a>'.format(r.OrganizationId, safe_str(r.OrganizationName)),
                safe_str(r.Division),
                safe_str(r.MemberType),
                format_date(r.InactiveDate)
            ])
        return self._table(headers, data, len(rows))

    def transactions(self, level='detail'):
        sql = """
            SELECT TOP 500 ts.PeopleId, p.Name2, ts.TranDate,
                   ts.TotPaid, ts.TotDue,
                   o.OrganizationId, o.OrganizationName,
                   d.Name as Division
            FROM TransactionSummary ts
            JOIN People p ON ts.PeopleId = p.PeopleId
            JOIN Organizations o ON ts.OrganizationId = o.OrganizationId
            LEFT JOIN Division d ON o.DivisionId = d.Id
            WHERE o.OrganizationId IN ({0})
            AND ts.TranDate >= '{1}' AND ts.TranDate <= '{2}'
            AND p.IsDeceased = 0 AND ISNULL(p.ArchivedFlag, 0) = 0
            ORDER BY ts.TranDate DESC
        """.format(self._org_id_list(), self.start_date, self.end_date)
        rows = list(q.QuerySql(sql))
        if not rows:
            return self._empty("No transactions this period.")
        total_paid = 0
        total_due = 0
        for r in rows:
            total_paid += r.TotPaid if r.TotPaid else 0
            total_due += r.TotDue if r.TotDue else 0
        summary = "<div style='margin-bottom:8px;font-weight:bold;'>Total Paid: ${0:,.2f} &bull; Total Due: ${1:,.2f} &bull; {2} transactions</div>".format(float(total_paid), float(total_due), len(rows))
        if level == 'summary':
            # Per-involvement breakdown with non-zero totals only
            org_totals = {}
            org_order = []
            for r in rows:
                oid = r.OrganizationId
                if oid not in org_totals:
                    org_totals[oid] = {'name': safe_str(r.OrganizationName), 'paid': 0, 'due': 0}
                    org_order.append(oid)
                org_totals[oid]['paid'] += r.TotPaid if r.TotPaid else 0
                org_totals[oid]['due'] += r.TotDue if r.TotDue else 0
            org_order.sort(key=lambda x: org_totals[x]['name'])
            summary += '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:6px;">'
            summary += '<thead><tr>'
            summary += '<th style="text-align:left;padding:4px 8px;border-bottom:1px solid #ddd;color:#888;font-size:11px;">INVOLVEMENT</th>'
            summary += '<th style="text-align:right;padding:4px 8px;border-bottom:1px solid #ddd;color:#888;font-size:11px;">PAID</th>'
            summary += '<th style="text-align:right;padding:4px 8px;border-bottom:1px solid #ddd;color:#888;font-size:11px;">DUE</th>'
            summary += '</tr></thead><tbody>'
            for oid in org_order:
                info = org_totals[oid]
                if info['paid'] == 0 and info['due'] == 0:
                    continue
                summary += '<tr>'
                summary += '<td style="padding:3px 8px;"><a href="/Organization/{0}" target="_blank">{1}</a></td>'.format(oid, info['name'])
                summary += '<td style="padding:3px 8px;text-align:right;font-weight:600;">${0:,.2f}</td>'.format(float(info['paid']))
                summary += '<td style="padding:3px 8px;text-align:right;font-weight:600;">${0:,.2f}</td>'.format(float(info['due']))
                summary += '</tr>'
            summary += '</tbody></table>'
            return summary
        headers = ['Person', 'Organization', 'Paid', 'Due', 'Date']
        data = []
        for r in rows:
            paid = r.TotPaid if r.TotPaid else 0
            due = r.TotDue if r.TotDue else 0
            data.append([
                '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(r.PeopleId, safe_str(r.Name2)),
                safe_str(r.OrganizationName),
                "${0:,.2f}".format(float(paid)),
                "${0:,.2f}".format(float(due)),
                format_date(r.TranDate)
            ])
        return summary + self._table(headers, data, len(rows))

    def new_people(self, level='detail'):
        sql = """
            SELECT DISTINCT TOP 200 p.PeopleId, p.Name2, p.EmailAddress,
                   p.CreatedDate, ms.Description as MemberStatus
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
            LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
            WHERE o.OrganizationId IN ({0})
            AND p.CreatedDate >= '{1}' AND p.CreatedDate <= '{2}'
            AND p.IsDeceased = 0 AND ISNULL(p.ArchivedFlag, 0) = 0
            ORDER BY p.CreatedDate DESC
        """.format(self._org_id_list(), self.start_date, self.end_date)
        rows = list(q.QuerySql(sql))
        if not rows:
            return self._empty("No new people records this period.")
        if level == 'summary':
            return self._summary_count(len(rows), "new person", "new people")
        headers = ['Person', 'Email', 'Status', 'Created']
        data = []
        for r in rows:
            data.append([
                '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(r.PeopleId, safe_str(r.Name2)),
                safe_str(r.EmailAddress),
                safe_str(r.MemberStatus),
                format_date(r.CreatedDate)
            ])
        return self._table(headers, data, len(rows))

    def new_baptisms(self, level='detail'):
        sql = """
            SELECT DISTINCT TOP 200 p.PeopleId, p.Name2, p.BaptismDate, p.Age
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
            WHERE o.OrganizationId IN ({0})
            AND p.BaptismDate >= '{1}' AND p.BaptismDate <= '{2}'
            AND p.IsDeceased = 0 AND ISNULL(p.ArchivedFlag, 0) = 0
            ORDER BY p.BaptismDate DESC
        """.format(self._org_id_list(), self.start_date, self.end_date)
        rows = list(q.QuerySql(sql))
        if not rows:
            return self._empty("No new baptisms this period.")
        if level == 'summary':
            return self._summary_count(len(rows), "new baptism", "new baptisms")
        headers = ['Person', 'Age', 'Baptism Date']
        data = []
        for r in rows:
            data.append([
                '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(r.PeopleId, safe_str(r.Name2)),
                str(r.Age) if r.Age else "",
                format_date(r.BaptismDate)
            ])
        return self._table(headers, data, len(rows))

    def new_church_members(self, level='detail'):
        sql = """
            SELECT DISTINCT TOP 200 p.PeopleId, p.Name2, p.JoinDate, p.Age,
                   ms.Description as MemberStatus
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            JOIN Organizations o ON om.OrganizationId = o.OrganizationId
            LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
            WHERE o.OrganizationId IN ({0})
            AND p.JoinDate >= '{1}' AND p.JoinDate <= '{2}'
            AND p.MemberStatusId = 10
            AND p.IsDeceased = 0 AND ISNULL(p.ArchivedFlag, 0) = 0
            ORDER BY p.JoinDate DESC
        """.format(self._org_id_list(), self.start_date, self.end_date)
        rows = list(q.QuerySql(sql))
        if not rows:
            return self._empty("No new church members this period.")
        if level == 'summary':
            return self._summary_count(len(rows), "new church member", "new church members")
        headers = ['Person', 'Age', 'Join Date']
        data = []
        for r in rows:
            data.append([
                '<a href="/Person2/{0}" target="_blank">{1}</a>'.format(r.PeopleId, safe_str(r.Name2)),
                str(r.Age) if r.Age else "",
                format_date(r.JoinDate)
            ])
        return self._table(headers, data, len(rows))

    def attendance_gaps(self, level='detail'):
        gap_weeks = self.config.get('attendanceGapWeeks', ATTENDANCE_GAP_WEEKS)
        sql = """
            WITH PersonAttend AS (
                SELECT
                    a.PeopleId,
                    p.Name2,
                    o.OrganizationId,
                    o.OrganizationName,
                    COUNT(CASE WHEN a.AttendanceFlag = 1
                        AND m.MeetingDate >= DATEADD(week, -{0}, GETDATE()) THEN 1 END) as RecentAttend,
                    COUNT(CASE WHEN a.AttendanceFlag = 1
                        AND m.MeetingDate >= DATEADD(week, -26, GETDATE())
                        AND m.MeetingDate < DATEADD(week, -{0}, GETDATE()) THEN 1 END) as PriorAttend,
                    MAX(CASE WHEN a.AttendanceFlag = 1 THEN m.MeetingDate END) as LastAttended
                FROM Attend a
                JOIN People p ON a.PeopleId = p.PeopleId
                JOIN Meetings m ON a.MeetingId = m.MeetingId
                JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                WHERE o.OrganizationId IN ({1})
                AND m.DidNotMeet = 0
                AND m.MeetingDate >= DATEADD(week, -26, GETDATE())
                AND p.IsDeceased = 0 AND ISNULL(p.ArchivedFlag, 0) = 0
                GROUP BY a.PeopleId, p.Name2, o.OrganizationId, o.OrganizationName
            ),
            MovedTo AS (
                SELECT pa2.PeopleId,
                       pa2.OrganizationId as MovedFromOrgId,
                       pa3.OrganizationName as NowAttending
                FROM PersonAttend pa2
                JOIN PersonAttend pa3 ON pa2.PeopleId = pa3.PeopleId
                    AND pa3.OrganizationId != pa2.OrganizationId
                    AND pa3.RecentAttend > 0
                WHERE pa2.PriorAttend >= 3 AND pa2.RecentAttend = 0
            ),
            OrgTotals AS (
                SELECT OrganizationId,
                       COUNT(*) as TotalPriorAttenders
                FROM PersonAttend
                WHERE PriorAttend >= 3
                GROUP BY OrganizationId
            )
            SELECT TOP 300 pa.PeopleId, pa.Name2,
                   pa.OrganizationId, pa.OrganizationName,
                   pa.PriorAttend, pa.LastAttended,
                   ot.TotalPriorAttenders,
                   mt.NowAttending
            FROM PersonAttend pa
            JOIN OrgTotals ot ON pa.OrganizationId = ot.OrganizationId
            LEFT JOIN MovedTo mt ON pa.PeopleId = mt.PeopleId
                AND pa.OrganizationId = mt.MovedFromOrgId
            WHERE pa.PriorAttend >= 3 AND pa.RecentAttend = 0
            ORDER BY pa.OrganizationName, pa.Name2
        """.format(gap_weeks, self._org_id_list())
        rows = list(q.QuerySql(sql))
        if not rows:
            return self._empty("No attendance gaps detected (people who previously attended but stopped).")
        if level == 'summary':
            return self._summary_count(len(rows), "person with attendance gap", "people with attendance gaps")

        # Fetch contact effort separately (lightweight query on small PeopleId set)
        contact_methods = get_contact_methods_config()
        keyword_id_to_code = {}
        all_keyword_ids = []
        for m in contact_methods:
            kid = m.get('keywordId')
            if kid:
                keyword_id_to_code[int(kid)] = m.get('code', '?')
                all_keyword_ids.append(int(kid))
        method_order = [m.get('code', '') for m in contact_methods]

        contact_map = {}  # PeopleId -> {methods: {code: count}, total: N, last_date: date}
        gap_people = {}  # PeopleId -> earliest LastAttended
        for r in rows:
            pid = r.PeopleId
            if pid not in gap_people or (r.LastAttended and (not gap_people[pid] or r.LastAttended < gap_people[pid])):
                gap_people[pid] = r.LastAttended
        if gap_people:
            pid_list = ','.join(str(p) for p in gap_people.keys())
            # Get total TaskNote counts per person (for "Other" calculation)
            total_sql = """
                SELECT tn.AboutPersonId as PeopleId,
                       COUNT(*) as TotalCount,
                       MAX(tn.CreatedDate) as LastContactDate
                FROM TaskNote tn WITH (NOLOCK)
                WHERE tn.AboutPersonId IN ({0})
                AND tn.CreatedDate >= DATEADD(week, -26, GETDATE())
                GROUP BY tn.AboutPersonId
            """.format(pid_list)
            for cr in q.QuerySql(total_sql):
                last_attended = gap_people.get(cr.PeopleId)
                if last_attended and cr.LastContactDate and cr.LastContactDate >= last_attended:
                    contact_map[cr.PeopleId] = {'methods': {}, 'total': int(cr.TotalCount) if cr.TotalCount else 0, 'last_date': cr.LastContactDate}

            # Get per-keyword counts (only if we have keyword IDs and people with contacts)
            if all_keyword_ids and contact_map:
                contacted_pids = ','.join(str(p) for p in contact_map.keys())
                kid_list = ','.join(str(k) for k in all_keyword_ids)
                method_sql = """
                    SELECT tn.AboutPersonId as PeopleId, tnk.KeywordId, COUNT(*) as Cnt
                    FROM TaskNote tn WITH (NOLOCK)
                    JOIN TaskNoteKeyword tnk ON tnk.TaskNoteId = tn.TaskNoteId
                    WHERE tn.AboutPersonId IN ({0})
                    AND tnk.KeywordId IN ({1})
                    AND tn.CreatedDate >= DATEADD(week, -26, GETDATE())
                    GROUP BY tn.AboutPersonId, tnk.KeywordId
                """.format(contacted_pids, kid_list)
                try:
                    for row in q.QuerySql(method_sql):
                        pid = row.PeopleId
                        if pid in contact_map:
                            code = keyword_id_to_code.get(int(row.KeywordId), '')
                            if code:
                                contact_map[pid]['methods'][code] = int(row.Cnt) if row.Cnt else 0
                except:
                    pass

        # Deduplicate rows (person may appear multiple times if moved to multiple orgs)
        # Collect NowAttending per person+org, then keep one row per person+org
        person_org_moved = {}  # (pid, oid) -> list of org names they moved to
        deduped = []
        seen = set()
        for r in rows:
            key = (r.PeopleId, r.OrganizationId)
            if r.NowAttending:
                if key not in person_org_moved:
                    person_org_moved[key] = []
                moved_name = safe_str(r.NowAttending)
                if moved_name not in person_org_moved[key]:
                    person_org_moved[key].append(moved_name)
            if key not in seen:
                seen.add(key)
                deduped.append(r)
        rows = deduped

        # Group by involvement
        involvements = {}
        inv_order = []
        for r in rows:
            oid = r.OrganizationId
            if oid not in involvements:
                total = r.TotalPriorAttenders if r.TotalPriorAttenders else 0
                involvements[oid] = {
                    'name': safe_str(r.OrganizationName),
                    'total': total,
                    'people': []
                }
                inv_order.append(oid)
            involvements[oid]['people'].append(r)

        # Calculate gap count and percentage in Python
        for oid in inv_order:
            inv = involvements[oid]
            inv['gap_count'] = len(inv['people'])
            inv['gap_pct'] = (float(inv['gap_count']) / inv['total'] * 100) if inv['total'] > 0 else 0.0

        # Sort by gap percentage descending (worst first)
        inv_order.sort(key=lambda x: involvements[x]['gap_pct'], reverse=True)

        # Build scope label from configured programs/divisions
        scope_parts = []
        for g in self.config.get('programDivGroups', []):
            pn = g.get('programName', '')
            dn = g.get('divisionName', '')
            if pn and dn:
                scope_parts.append('{0} &gt; {1}'.format(safe_str(pn), safe_str(dn)))
            elif pn:
                scope_parts.append(safe_str(pn))
        scope_label = ', '.join(scope_parts) if scope_parts else 'selected programs'

        html = '<div style="margin-bottom:10px;padding:10px 14px;background:#f0f7ff;border-left:3px solid #3498db;border-radius:0 4px 4px 0;font-size:12px;color:#555;line-height:1.5;">'
        html += 'People who attended <strong>3+ times</strong> in the prior 6 months but have <strong>not attended this involvement</strong> in the last <strong>{0} week{1}</strong>. '.format(gap_weeks, 's' if gap_weeks != 1 else '')
        method_legend = ', '.join(['<span style="color:#27ae60;font-weight:700;">{0}</span>={1}'.format(m.get('code',''), m.get('label','')) for m in contact_methods]) if contact_methods else ''
        if method_legend:
            method_legend += ', <span style="color:#888;font-weight:700;">O</span>=Other'
        contact_legend = 'Contacted shows task/note activity since last attended ({0}).'.format(method_legend) if method_legend else 'Contacted shows task/note count since last attended (<span style="color:#888;font-weight:700;">O</span>=Other).'
        html += '<span style="color:#2980b9;">Moved to</span> = now attending another monitored involvement. <span style="color:#e74c3c;">Disengaged</span> = not attending any monitored involvement ({0}). {1} Sorted by lapse rate (worst first).'.format(scope_label, contact_legend)
        html += '</div>'
        html += '<div style="margin-bottom:8px;color:#666;font-size:13px;">{0} people across {1} involvements</div>'.format(len(rows), len(inv_order))

        for oid in inv_order:
            inv = involvements[oid]
            pct = inv['gap_pct']
            # Color code: red > 50%, orange > 25%, green otherwise
            if pct > 50:
                bar_color = '#e74c3c'
            elif pct > 25:
                bar_color = '#f39c12'
            else:
                bar_color = '#27ae60'

            html += '<div style="margin-bottom:16px;border:1px solid #eee;border-radius:6px;overflow:hidden;">'
            # Involvement header with gap bar
            html += '<div style="padding:10px 14px;background:#f8f9fa;border-bottom:1px solid #eee;">'
            html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">'
            html += '<a href="/Organization/{0}" target="_blank" style="font-weight:600;font-size:14px;">{1}</a>'.format(oid, inv['name'])
            html += '<span style="font-size:12px;color:#666;">{0} of {1} lapsed ({2:.0f}%)</span>'.format(inv['gap_count'], inv['total'], float(pct))
            html += '</div>'
            # Progress bar
            html += '<div style="background:#e9ecef;border-radius:4px;height:8px;overflow:hidden;">'
            html += '<div style="width:{0:.0f}%;height:100%;background:{1};border-radius:4px;"></div>'.format(float(min(pct, 100)), bar_color)
            html += '</div>'
            html += '</div>'

            # People in this involvement
            html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
            html += '<thead><tr>'
            html += '<th style="text-align:left;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">PERSON</th>'
            html += '<th style="text-align:right;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">PRIOR ATTEND</th>'
            html += '<th style="text-align:right;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">LAST ATTENDED</th>'
            html += '<th style="text-align:left;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">STATUS</th>'
            html += '<th style="text-align:center;padding:6px 14px;border-bottom:1px solid #eee;color:#888;font-size:11px;font-weight:600;">CONTACTED</th>'
            html += '</tr></thead><tbody>'
            for i, r in enumerate(inv['people']):
                bg = '#fff' if i % 2 == 0 else '#f8f9fa'
                moved = person_org_moved.get((r.PeopleId, r.OrganizationId), [])
                html += '<tr style="background:{0};">'.format(bg)
                html += '<td style="padding:5px 14px;"><a href="/Person2/{0}" target="_blank">{1}</a></td>'.format(r.PeopleId, safe_str(r.Name2))
                html += '<td style="padding:5px 14px;text-align:right;">{0}</td>'.format(r.PriorAttend)
                html += '<td style="padding:5px 14px;text-align:right;">{0}</td>'.format(format_date(r.LastAttended))
                if moved:
                    html += '<td style="padding:5px 14px;font-size:11px;color:#2980b9;">Moved to: {0}</td>'.format(', '.join(moved))
                else:
                    html += '<td style="padding:5px 14px;font-size:11px;color:#e74c3c;">Disengaged</td>'
                ce = contact_map.get(r.PeopleId, {})
                total = ce.get('total', 0)
                if total > 0:
                    methods = ce.get('methods', {})
                    known_sum = sum(methods.values())
                    other_count = total - known_sum
                    parts = []
                    for code in method_order:
                        cnt = methods.get(code, 0)
                        if cnt > 0:
                            parts.append('<span style="color:#27ae60;font-weight:700;">{0}</span><sup style="font-size:9px;">{1}</sup>'.format(code, cnt))
                    if other_count > 0:
                        parts.append('<span style="color:#888;font-weight:700;">O</span><sup style="font-size:9px;">{0}</sup>'.format(other_count))
                    last_contact = format_date(ce.get('last_date')) if ce.get('last_date') else ''
                    html += '<td style="padding:5px 14px;text-align:center;font-size:11px;">'
                    html += '&nbsp;&nbsp;'.join(parts) if parts else str(total)
                    if last_contact:
                        html += '<br/><span style="color:#888;font-size:10px;">{0}</span>'.format(last_contact)
                    html += '</td>'
                else:
                    html += '<td style="padding:5px 14px;text-align:center;font-size:11px;color:#ccc;">&mdash;</td>'
                html += '</tr>'
            html += '</tbody></table></div>'

        return html

    EMPTY_SENTINEL = '<!--EMPTY_SECTION-->'

    def _empty(self, msg):
        return self.EMPTY_SENTINEL + '<p style="color:#888;font-style:italic;margin:8px 0;">{0}</p>'.format(msg)

    def _org_summary(self, rows, total, singular, plural):
        """Summary with per-organization breakdown."""
        label = singular if total == 1 else plural
        color = '#27ae60' if total > 0 else '#888'
        html = '<div style="font-size:20px;font-weight:bold;color:{0};padding:8px 0;">{1} {2}</div>'.format(color, total, label)
        # Group by organization
        org_counts = {}
        org_order = []
        for r in rows:
            oid = r.OrganizationId
            if oid not in org_counts:
                org_counts[oid] = {'name': safe_str(r.OrganizationName), 'count': 0}
                org_order.append(oid)
            org_counts[oid]['count'] += 1
        org_order.sort(key=lambda x: org_counts[x]['name'])
        html += '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:6px;">'
        for oid in org_order:
            info = org_counts[oid]
            html += '<tr><td style="padding:3px 8px;"><a href="/Organization/{0}" target="_blank">{1}</a></td><td style="padding:3px 8px;text-align:right;font-weight:600;">{2}</td></tr>'.format(oid, info['name'], info['count'])
        html += '</table>'
        return html

    def _summary_count(self, count, singular, plural):
        label = singular if count == 1 else plural
        color = '#27ae60' if count > 0 else '#888'
        return '<div style="font-size:20px;font-weight:bold;color:{0};padding:8px 0;">{1}</div><div style="color:#666;font-size:13px;">{2}</div>'.format(color, count, label)

    def _table(self, headers, data, count):
        html = []
        html.append('<div style="margin-bottom:4px;color:#666;font-size:13px;">{0} record{1}</div>'.format(count, 's' if count != 1 else ''))
        html.append('<table style="width:100%;border-collapse:collapse;font-size:13px;">')
        html.append('<thead><tr>')
        for h in headers:
            html.append('<th style="text-align:left;padding:6px 8px;border-bottom:2px solid #ddd;background:#f8f9fa;">{0}</th>'.format(h))
        html.append('</tr></thead><tbody>')
        for i, row in enumerate(data):
            bg = '#fff' if i % 2 == 0 else '#f8f9fa'
            html.append('<tr style="background:{0};">'.format(bg))
            for cell in row:
                html.append('<td style="padding:5px 8px;border-bottom:1px solid #eee;">{0}</td>'.format(cell if cell else ''))
            html.append('</tr>')
        html.append('</tbody></table>')
        return ''.join(html)

    def build_report(self, sections):
        config_name = self.config.get('name', 'Program Pulse')
        prog_names = []
        for g in self.config.get('programDivGroups', []):
            pn = g.get('programName', '')
            dn = g.get('divisionName', '')
            if dn:
                prog_names.append("{0} > {1}".format(pn, dn))
            else:
                prog_names.append(pn)

        html = []
        html.append('<div style="font-family:Segoe UI,Arial,sans-serif;max-width:900px;margin:0 auto;">')

        # Header
        html.append('<div style="background:#2c3e50;color:#fff;padding:16px 24px;border-radius:6px 6px 0 0;">')
        html.append('<h2 style="margin:0;font-size:22px;">{0}</h2>'.format(safe_str(config_name)))
        html.append('<div style="margin-top:6px;font-size:13px;opacity:0.85;">')
        html.append('{0} &mdash; {1}'.format(format_date(self.start_date), format_date(self.end_date)))
        if prog_names:
            html.append(' &bull; {0}'.format(', '.join(safe_str(n) for n in prog_names)))
        html.append('</div></div>')

        # Summary bar
        html.append('<div style="background:#ecf0f1;padding:12px 24px;border-bottom:1px solid #ddd;">')
        html.append('<strong>{0}</strong> sections &bull; <strong>{1}</strong> organizations monitored'.format(
            len(sections), len(self.org_ids)))
        html.append('</div>')

        # Sections
        for title, content in sections:
            html.append('<div style="margin:0;border:1px solid #e0e0e0;border-top:none;">')
            html.append('<div style="background:#3498db;color:#fff;padding:10px 24px;font-weight:bold;font-size:15px;">{0}</div>'.format(title))
            html.append('<div style="padding:12px 24px;">{0}</div>'.format(content))
            html.append('</div>')

        # Footer
        html.append('<div style="padding:12px 24px;color:#999;font-size:11px;text-align:center;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 6px 6px;">')
        html.append('Generated {0} by Program Pulse'.format(datetime.datetime.now().strftime("%m/%d/%Y %I:%M %p")))
        html.append('</div>')

        html.append('</div>')
        return ''.join(html)


# ============================================================
# EMAIL SENDER
# ============================================================

def send_report(config, html_content):
    recipients = config.get('recipients', [])
    # Use configured sender from global settings, fallback to current user
    settings = load_settings()
    configured_sender = settings.get('sender')
    if configured_sender and configured_sender.get('peopleId'):
        sender_id = int(configured_sender['peopleId'])
        sender = model.GetPerson(sender_id)
    else:
        sender_id = model.UserPeopleId
        sender = model.GetPerson(sender_id)
    sender_email = sender.EmailAddress if sender else ""
    sender_name = sender.Name if sender else ""
    config_name = config.get('name', 'Program Pulse')
    subject = "Program Pulse: {0}".format(safe_str(config_name))

    sent_count = 0
    errors = []

    for recip in recipients:
        pid = recip.get('peopleId')
        if not pid:
            continue
        try:
            person = model.GetPerson(int(pid))
            if person and person.EmailAddress:
                query = "PeopleId={0}".format(int(pid))
                model.Email(query, sender_id, sender_email, sender_name, subject, html_content, "")
                sent_count += 1
        except Exception as e:
            errors.append("PeopleId {0}: {1}".format(pid, safe_str(e)))

    return sent_count, errors


# ============================================================
# AJAX HANDLERS
# ============================================================

def handle_ajax():
    action = get_form_data('action', '')
    response = {'success': False, 'message': 'Unknown action'}

    try:
        if action == 'load_configs':
            configs = load_configs()
            response = {'success': True, 'configs': configs}

        elif action == 'save_config':
            config_json = get_form_data('config_data', '{}')
            config = json.loads(config_json)
            configs = load_configs()

            config_id = config.get('id', '')
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

            if config_id:
                found = False
                for i, c in enumerate(configs):
                    if c.get('id') == config_id:
                        config['updatedAt'] = now_str
                        configs[i] = config
                        found = True
                        break
                if not found:
                    config['createdAt'] = now_str
                    config['updatedAt'] = now_str
                    configs.append(config)
            else:
                config['id'] = generate_id()
                config['createdAt'] = now_str
                config['updatedAt'] = now_str
                configs.append(config)

            save_configs(configs)
            response = {'success': True, 'message': 'Configuration saved', 'config': config}

        elif action == 'delete_config':
            config_id = get_form_data('config_id', '')
            configs = load_configs()
            configs = [c for c in configs if c.get('id') != config_id]
            save_configs(configs)
            response = {'success': True, 'message': 'Configuration deleted'}

        elif action == 'get_programs':
            sql = """
                SELECT p.Id, p.Name,
                    (SELECT COUNT(*) FROM Organizations o
                     JOIN Division d2 ON o.DivisionId = d2.Id
                     WHERE d2.ProgId = p.Id AND o.OrganizationStatusId = 30) as OrgCount
                FROM Program p
                ORDER BY p.Name
            """
            programs = []
            for row in q.QuerySql(sql):
                programs.append({
                    'id': row.Id,
                    'name': safe_str(row.Name),
                    'orgCount': row.OrgCount
                })
            response = {'success': True, 'programs': programs}

        elif action == 'get_divisions':
            prog_id = get_form_data('program_id', '0')
            sql = """
                SELECT DISTINCT d.Id, d.Name,
                    (SELECT COUNT(*) FROM Organizations o
                     WHERE o.DivisionId = d.Id AND o.OrganizationStatusId = 30) as OrgCount
                FROM Division d
                JOIN Organizations o ON o.DivisionId = d.Id
                WHERE d.ProgId = {0}
                AND o.OrganizationStatusId = 30
                ORDER BY d.Name
            """.format(safe_int(prog_id))
            divisions = []
            for row in q.QuerySql(sql):
                divisions.append({
                    'id': row.Id,
                    'name': safe_str(row.Name),
                    'orgCount': row.OrgCount
                })
            response = {'success': True, 'divisions': divisions}

        elif action == 'search_people':
            search_term = get_form_data('search_term', '')
            if len(search_term) < 2:
                response = {'success': True, 'people': []}
            else:
                sql = """
                    SELECT TOP 20 p.PeopleId, p.Name2, p.EmailAddress
                    FROM People p
                    WHERE p.IsDeceased = 0
                    AND (p.Name2 LIKE '%{0}%' OR p.Name LIKE '%{0}%' OR p.EmailAddress LIKE '%{0}%')
                    ORDER BY p.Name2
                """.format(search_term.replace("'", "''"))
                people = []
                for row in q.QuerySql(sql):
                    people.append({
                        'peopleId': row.PeopleId,
                        'name': safe_str(row.Name2),
                        'email': safe_str(row.EmailAddress)
                    })
                response = {'success': True, 'people': people}

        elif action == 'generate_preview':
            config_json = get_form_data('config_data', '{}')
            config = json.loads(config_json)
            lookback = safe_int(config.get('lookbackDays', DEFAULT_LOOKBACK_DAYS), DEFAULT_LOOKBACK_DAYS)
            end_date = datetime.datetime.now().strftime('%Y-%m-%d')
            start_date = (datetime.datetime.now() - datetime.timedelta(days=lookback)).strftime('%Y-%m-%d')

            gen = ReportGenerator(config, start_date, end_date)
            html = gen.generate()
            response = {'success': True, 'html': html}

        elif action == 'send_report':
            config_id = get_form_data('config_id', '')
            configs = load_configs()
            config = None
            for c in configs:
                if c.get('id') == config_id:
                    config = c
                    break
            if not config:
                response = {'success': False, 'message': 'Configuration not found'}
            else:
                lookback = safe_int(config.get('lookbackDays', DEFAULT_LOOKBACK_DAYS), DEFAULT_LOOKBACK_DAYS)
                end_date = datetime.datetime.now().strftime('%Y-%m-%d')
                start_date = (datetime.datetime.now() - datetime.timedelta(days=lookback)).strftime('%Y-%m-%d')

                gen = ReportGenerator(config, start_date, end_date)
                html = gen.generate()

                sent, errors = send_report(config, html)
                append_log({
                    'timestamp': datetime.datetime.now().isoformat(),
                    'configId': config_id,
                    'configName': config.get('name', ''),
                    'sent': sent,
                    'errors': errors,
                    'source': 'manual'
                })
                response = {'success': True, 'sent': sent, 'errors': errors}

        elif action == 'get_log':
            logs = load_log()
            response = {'success': True, 'logs': logs[-20:]}

        elif action == 'load_settings':
            settings = load_settings()
            response = {'success': True, 'settings': settings}

        elif action == 'save_settings':
            settings_json = get_form_data('settings_data', '{}')
            settings = json.loads(settings_json)
            save_settings(settings)
            response = {'success': True, 'message': 'Settings saved'}

        elif action == 'get_keywords':
            sql = """
                SELECT KeywordId, Description
                FROM dbo.Keyword
                WHERE IsActive = 1
                ORDER BY Description
            """
            keywords = []
            for row in q.QuerySql(sql):
                keywords.append({'keywordId': int(row.KeywordId), 'description': str(row.Description)})
            response = {'success': True, 'keywords': keywords}

    except Exception as e:
        response = {'success': False, 'message': safe_str(e)}

    print json.dumps(response)


# ============================================================
# BATCH HANDLER
# ============================================================

def handle_batch():
    configs = load_configs()
    today = datetime.datetime.now()
    day_of_week = today.weekday()

    for config in configs:
        if not config.get('enabled', False):
            continue

        sched_day = config.get('scheduledDay', -1)
        if sched_day != day_of_week:
            continue

        try:
            lookback = safe_int(config.get('lookbackDays', DEFAULT_LOOKBACK_DAYS), DEFAULT_LOOKBACK_DAYS)
            end_date = today.strftime('%Y-%m-%d')
            start_date = (today - datetime.timedelta(days=lookback)).strftime('%Y-%m-%d')

            gen = ReportGenerator(config, start_date, end_date)
            html = gen.generate()
            sent, errors = send_report(config, html)

            append_log({
                'timestamp': today.isoformat(),
                'configId': config.get('id', ''),
                'configName': config.get('name', ''),
                'sent': sent,
                'errors': errors,
                'source': 'batch'
            })
        except Exception as e:
            append_log({
                'timestamp': today.isoformat(),
                'configId': config.get('id', ''),
                'configName': config.get('name', ''),
                'sent': 0,
                'errors': [safe_str(e)],
                'source': 'batch'
            })


# ============================================================
# UI HTML
# ============================================================

def generate_ui():
    return '''
<script>
if (window.location.pathname.indexOf('/PyScript/') > -1) {
    window.location.href = window.location.pathname.replace('/PyScript/', '/PyScriptForm/') + window.location.search;
}
</script>
<style>
.pp-container { font-family: 'Segoe UI', Tahoma, sans-serif; max-width: 1100px; margin: 0 auto; }
.pp-card { background: #fff; border: 1px solid #ddd; border-radius: 6px; margin-bottom: 16px; }
.pp-card-header { background: #f8f9fa; padding: 12px 20px; border-bottom: 1px solid #ddd; border-radius: 6px 6px 0 0; font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
.pp-card-body { padding: 16px 20px; }
.pp-btn { display: inline-block; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: 500; text-decoration: none; }
.pp-btn-primary { background: #3498db; color: #fff; }
.pp-btn-primary:hover { background: #2980b9; }
.pp-btn-success { background: #27ae60; color: #fff; }
.pp-btn-success:hover { background: #219a52; }
.pp-btn-danger { background: #e74c3c; color: #fff; }
.pp-btn-danger:hover { background: #c0392b; }
.pp-btn-outline { background: #fff; color: #333; border: 1px solid #ccc; }
.pp-btn-outline:hover { background: #f0f0f0; }
.pp-btn-sm { padding: 4px 10px; font-size: 12px; }
.pp-form-group { margin-bottom: 14px; }
.pp-form-group label { display: block; font-weight: 600; margin-bottom: 4px; font-size: 13px; color: #555; }
.pp-input, .pp-select { width: 100%; padding: 8px 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; box-sizing: border-box; }
.pp-input:focus, .pp-select:focus { border-color: #3498db; outline: none; box-shadow: 0 0 0 2px rgba(52,152,219,0.15); }
.pp-row { display: flex; gap: 16px; flex-wrap: wrap; }
.pp-col-half { flex: 1; min-width: 280px; }
.pp-tag { display: inline-flex; align-items: center; background: #e8f4fd; color: #2980b9; padding: 4px 10px; border-radius: 16px; font-size: 12px; margin: 2px 4px 2px 0; }
.pp-tag .pp-tag-remove { margin-left: 6px; cursor: pointer; color: #c0392b; font-weight: bold; }
.pp-tag .pp-tag-remove:hover { color: #e74c3c; }
.pp-checkbox-group { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 6px; }
.pp-checkbox-group label { display: flex; align-items: center; gap: 6px; font-weight: normal; cursor: pointer; padding: 6px 8px; border-radius: 4px; }
.pp-checkbox-group label:hover { background: #f0f7ff; }
.pp-config-row { display: flex; justify-content: space-between; align-items: center; padding: 10px 16px; border-bottom: 1px solid #eee; }
.pp-config-row:last-child { border-bottom: none; }
.pp-config-row:hover { background: #f8f9fa; }
.pp-config-name { font-weight: 600; }
.pp-config-meta { font-size: 12px; color: #888; margin-top: 2px; }
.pp-config-actions { display: flex; gap: 6px; }
.pp-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
.pp-badge-on { background: #d5f5e3; color: #27ae60; }
.pp-badge-off { background: #fadbd8; color: #e74c3c; }
.pp-tabs { display: flex; gap: 0; border-bottom: 2px solid #ddd; margin-bottom: 16px; }
.pp-tab { padding: 10px 20px; cursor: pointer; font-weight: 500; color: #888; border-bottom: 2px solid transparent; margin-bottom: -2px; }
.pp-tab:hover { color: #333; }
.pp-tab.active { color: #3498db; border-bottom-color: #3498db; }
.pp-tab-content { display: none; }
.pp-tab-content.active { display: block; }
.pp-preview-frame { border: 1px solid #ddd; border-radius: 4px; padding: 16px; min-height: 200px; background: #fafafa; }
.pp-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000; justify-content: center; align-items: flex-start; padding-top: 40px; }
.pp-overlay.active { display: flex; }
.pp-modal { background: #fff; border-radius: 8px; width: 90%; max-width: 800px; max-height: 85vh; overflow-y: auto; box-shadow: 0 8px 32px rgba(0,0,0,0.2); }
.pp-modal-header { padding: 16px 24px; border-bottom: 1px solid #ddd; display: flex; justify-content: space-between; align-items: center; position: sticky; top: 0; background: #fff; z-index: 1; border-radius: 8px 8px 0 0; }
.pp-modal-body { padding: 20px 24px; }
.pp-modal-footer { padding: 12px 24px; border-top: 1px solid #ddd; text-align: right; display: flex; gap: 8px; justify-content: flex-end; position: sticky; bottom: 0; background: #fff; border-radius: 0 0 8px 8px; }
.pp-search-results { max-height: 200px; overflow-y: auto; border: 1px solid #ddd; border-top: none; border-radius: 0 0 4px 4px; }
.pp-search-item { padding: 8px 12px; cursor: pointer; border-bottom: 1px solid #f0f0f0; font-size: 13px; }
.pp-search-item:hover { background: #e8f4fd; }
.pp-log-entry { padding: 8px 12px; border-bottom: 1px solid #eee; font-size: 12px; }
.pp-empty { text-align: center; padding: 40px 20px; color: #999; }
.pp-spinner { display: inline-block; width: 16px; height: 16px; border: 2px solid #ddd; border-top-color: #3498db; border-radius: 50%; animation: pp-spin 0.6s linear infinite; }
@keyframes pp-spin { to { transform: rotate(360deg); } }
.pp-day-selector { display: flex; gap: 4px; flex-wrap: wrap; }
.pp-day-btn { padding: 6px 12px; border: 1px solid #ccc; border-radius: 4px; cursor: pointer; font-size: 12px; background: #fff; }
.pp-day-btn.selected { background: #3498db; color: #fff; border-color: #3498db; }
.pp-help-toggle { display: inline-block; width: 16px; height: 16px; line-height: 16px; text-align: center; background: #e8f4fd; color: #3498db; border-radius: 50%; font-size: 11px; font-weight: 700; cursor: pointer; margin-left: 4px; vertical-align: middle; user-select: none; }
.pp-help-toggle:hover { background: #3498db; color: #fff; }
.pp-help-text { display: none; background: #f0f7ff; border-left: 3px solid #3498db; padding: 8px 12px; margin: 6px 0 10px 0; font-size: 12px; color: #555; line-height: 1.5; border-radius: 0 4px 4px 0; }
.pp-help-text.active { display: block; }
.pp-sender-banner { background: #fff3cd; border: 1px solid #ffc107; border-radius: 6px; padding: 12px 16px; margin-bottom: 16px; display: flex; align-items: center; gap: 12px; font-size: 13px; color: #856404; }
.pp-sender-banner a { color: #3498db; font-weight: 600; cursor: pointer; text-decoration: underline; }
.pp-tutorial { line-height: 1.7; color: #444; }
.pp-tutorial h3 { color: #2c3e50; margin: 24px 0 8px 0; font-size: 16px; border-bottom: 1px solid #eee; padding-bottom: 6px; }
.pp-tutorial h3:first-child { margin-top: 0; }
.pp-tutorial ol, .pp-tutorial ul { margin: 8px 0 8px 20px; }
.pp-tutorial li { margin-bottom: 6px; }
.pp-tutorial code { background: #f4f4f4; padding: 2px 6px; border-radius: 3px; font-size: 12px; }
.pp-tutorial .pp-step { background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 6px; padding: 14px 18px; margin-bottom: 12px; }
.pp-tutorial .pp-step-num { display: inline-block; width: 24px; height: 24px; line-height: 24px; text-align: center; background: #3498db; color: #fff; border-radius: 50%; font-weight: 700; font-size: 13px; margin-right: 8px; }
</style>

<div class="pp-container" id="ppApp">

    <div class="pp-tabs">
        <div class="pp-tab active" onclick="ppSwitchTab('configs')">Configurations</div>
        <div class="pp-tab" onclick="ppSwitchTab('preview')">Preview</div>
        <div class="pp-tab" onclick="ppSwitchTab('log')">Send History</div>
        <div class="pp-tab" onclick="ppSwitchTab('settings')">Settings</div>
        <div class="pp-tab" onclick="ppSwitchTab('help')">Help</div>
    </div>

    <!-- CONFIGS TAB -->
    <div class="pp-tab-content active" id="ppTab_configs">
        <div class="pp-sender-banner" id="ppSenderBanner" style="display:none;">
            <span style="font-size:18px;">&#9888;</span>
            <div>
                <strong>No default sender configured.</strong> Reports will be sent from whoever triggers them.
                <a onclick="ppSwitchTab('settings')">Set a default sender in Settings</a> so reports always come from a consistent address.
            </div>
        </div>
        <div class="pp-card">
            <div class="pp-card-header">
                <span>Report Configurations</span>
                <button class="pp-btn pp-btn-primary" onclick="ppOpenEditor(null)">+ New Configuration</button>
            </div>
            <div class="pp-card-body" id="ppConfigList">
                <div class="pp-empty">Loading configurations...</div>
            </div>
        </div>
    </div>

    <!-- PREVIEW TAB -->
    <div class="pp-tab-content" id="ppTab_preview">
        <div class="pp-card">
            <div class="pp-card-header">
                <span>Report Preview</span>
                <div style="display:flex;gap:8px;">
                    <select class="pp-select" id="ppPreviewConfig" style="width:250px;" onchange="ppLoadPreview()">
                        <option value="">-- Select Configuration --</option>
                    </select>
                    <button class="pp-btn pp-btn-primary" onclick="ppLoadPreview()">Generate Preview</button>
                </div>
            </div>
            <div class="pp-card-body">
                <div class="pp-preview-frame" id="ppPreviewArea">
                    <div class="pp-empty">Select a configuration and click Generate Preview</div>
                </div>
            </div>
        </div>
    </div>

    <!-- LOG TAB -->
    <div class="pp-tab-content" id="ppTab_log">
        <div class="pp-card">
            <div class="pp-card-header">
                <span>Send History</span>
                <button class="pp-btn pp-btn-outline pp-btn-sm" onclick="ppLoadLog()">Refresh</button>
            </div>
            <div class="pp-card-body" id="ppLogArea">
                <div class="pp-empty">Loading...</div>
            </div>
        </div>
    </div>

    <!-- SETTINGS TAB -->
    <div class="pp-tab-content" id="ppTab_settings">
        <div class="pp-card">
            <div class="pp-card-header">
                <span>Global Settings</span>
            </div>
            <div class="pp-card-body">
                <div style="background:#e8f4fd;border:1px solid #b8daff;border-radius:4px;padding:12px 16px;margin-bottom:16px;font-size:13px;color:#004085;">
                    Configure the contact methods (P, E, T, V, M, etc.) and their associated TouchPoint keywords. These are used in the Attendance Gaps section to show contact effort breakdown. <strong style="color:#888;">O=Other</strong> is a built-in catch-all that automatically counts any task/note activity not matching a defined method below.
                </div>
                <h4 style="margin:0 0 12px 0;font-size:14px;">Default Sender</h4>
                <div style="margin-bottom:16px;">
                    <div style="font-size:12px;color:#666;margin-bottom:6px;">Reports will be sent from this person. If not set, uses the currently logged-in user.</div>
                    <div style="display:flex;gap:8px;align-items:center;">
                        <input class="pp-input" id="ppSenderSearch" placeholder="Search by name..." style="width:300px;" oninput="ppSearchSender(this.value)"/>
                        <span id="ppSenderDisplay" style="font-size:13px;color:#27ae60;font-weight:600;"></span>
                    </div>
                    <div class="pp-search-results" id="ppSenderResults" style="display:none;max-width:300px;"></div>
                </div>
                <hr style="border:none;border-top:1px solid #eee;margin:16px 0;"/>
                <h4 style="margin:0 0 12px 0;font-size:14px;">Contact Methods</h4>
                <table style="width:100%;border-collapse:collapse;font-size:13px;" id="ppContactMethodsTable">
                    <thead>
                        <tr>
                            <th style="text-align:left;padding:8px;border-bottom:2px solid #ddd;width:60px;">Code</th>
                            <th style="text-align:left;padding:8px;border-bottom:2px solid #ddd;width:140px;">Label</th>
                            <th style="text-align:left;padding:8px;border-bottom:2px solid #ddd;">TouchPoint Keyword</th>
                            <th style="width:80px;padding:8px;border-bottom:2px solid #ddd;"></th>
                        </tr>
                    </thead>
                    <tbody id="ppContactMethodRows"></tbody>
                </table>
                <div style="margin-top:10px;">
                    <button class="pp-btn pp-btn-outline pp-btn-sm" onclick="ppAddContactMethod()">+ Add Contact Method</button>
                </div>
                <div style="margin-top:16px;display:flex;gap:8px;">
                    <button class="pp-btn pp-btn-success" onclick="ppSaveSettings()">Save Settings</button>
                    <span id="ppSettingsStatus" style="font-size:12px;color:#27ae60;line-height:32px;"></span>
                </div>
            </div>
        </div>
    </div>

    <!-- HELP TAB -->
    <div class="pp-tab-content" id="ppTab_help">
        <div class="pp-card">
            <div class="pp-card-header">
                <span>Getting Started with Program Pulse</span>
            </div>
            <div class="pp-card-body pp-tutorial">

                <h3>What is Program Pulse?</h3>
                <p>Program Pulse generates scheduled digest reports summarizing activity across your programs and divisions. It tracks new enrollments, dropped members, attendance gaps, transactions, baptisms, new members, and more &mdash; then emails a summary to the people you choose.</p>

                <h3>Quick Start</h3>

                <div class="pp-step">
                    <span class="pp-step-num">1</span>
                    <strong>Set a Default Sender</strong>
                    <p style="margin:6px 0 0 32px;">Go to the <strong>Settings</strong> tab and search for the person whose email address should appear as the "From" on all reports. This ensures consistency regardless of who triggers the report.</p>
                </div>

                <div class="pp-step">
                    <span class="pp-step-num">2</span>
                    <strong>Create a Configuration</strong>
                    <p style="margin:6px 0 0 32px;">Click <strong>+ New Configuration</strong> on the Configurations tab. Give it a name (e.g., "Adults Weekly Summary"), select the programs/divisions to monitor, choose which report sections to include, and add recipients.</p>
                </div>

                <div class="pp-step">
                    <span class="pp-step-num">3</span>
                    <strong>Preview Your Report</strong>
                    <p style="margin:6px 0 0 32px;">Go to the <strong>Preview</strong> tab, select your configuration, and click <strong>Generate Preview</strong> to see exactly what the email will look like.</p>
                </div>

                <div class="pp-step">
                    <span class="pp-step-num">4</span>
                    <strong>Schedule or Send</strong>
                    <p style="margin:6px 0 0 32px;">In your configuration, pick a <strong>Scheduled Send Day</strong> and check <strong>Enable automatic sending</strong>. The morning batch will send it on that day each week. You can also click <strong>Send Now</strong> from the Configurations tab at any time.</p>
                </div>

                <h3>Report Sections Explained</h3>
                <ul>
                    <li><strong>New Enrollments</strong> &mdash; People who joined an organization during the lookback period.</li>
                    <li><strong>Dropped Members</strong> &mdash; People whose membership was inactivated during the period.</li>
                    <li><strong>New Involvements Created</strong> &mdash; Organizations created within your monitored programs during the period.</li>
                    <li><strong>Stale Involvements</strong> &mdash; Active organizations with no meeting activity beyond the stale threshold (default: 30 days).</li>
                    <li><strong>Transactions</strong> &mdash; Payment activity (fees paid, balances due) within monitored organizations.</li>
                    <li><strong>New People</strong> &mdash; Person records created during the period who are connected to your monitored organizations.</li>
                    <li><strong>New Baptisms</strong> &mdash; People with baptism dates during the period.</li>
                    <li><strong>New Church Members</strong> &mdash; People with join dates during the period whose member status is "Member."</li>
                    <li><strong>Attendance Gaps</strong> &mdash; People who attended 3+ times in the prior 6 months but have stopped attending in the recent gap window. Shows whether they moved to another monitored involvement or are disengaged, plus contact effort from task/note records.</li>
                </ul>

                <h3>Section Detail Levels</h3>
                <p>Each section can be set to one of three levels:</p>
                <ul>
                    <li><strong>Detail</strong> &mdash; Full table with individual records and links.</li>
                    <li><strong>Summary</strong> &mdash; Counts and per-organization breakdown only.</li>
                    <li><strong>None</strong> &mdash; Section is excluded from the report entirely.</li>
                </ul>

                <h3>Contact Methods (Settings)</h3>
                <p>The Attendance Gaps section can show a breakdown of outreach efforts per person. To enable this:</p>
                <ol>
                    <li>Go to <strong>Settings</strong> and add contact methods (e.g., <code>P</code> = Phone Call, <code>E</code> = Email, <code>V</code> = Visit).</li>
                    <li>Link each method to a TouchPoint <strong>Keyword</strong> used on task/note records.</li>
                    <li>The report will show method-coded badges next to each person with an attendance gap, along with an <code>O</code> (Other) catch-all for unmatched task/note activity.</li>
                </ol>

                <h3>Installation</h3>

                <div class="pp-step">
                    <span class="pp-step-num">1</span>
                    <strong>Add the Script</strong>
                    <p style="margin:6px 0 0 32px;">Go to <strong>Admin &gt; Advanced &gt; Special Content &gt; Python</strong>. Click <strong>Add New</strong>, name it <code>TPxi_ProgramPulse</code>, paste the full script code, and save.</p>
                </div>

                <div class="pp-step">
                    <span class="pp-step-num">2</span>
                    <strong>Access the Dashboard</strong>
                    <p style="margin:6px 0 0 32px;">Navigate to <code>/PyScript/TPxi_ProgramPulse</code> on your TouchPoint site. You should see the Program Pulse dashboard.</p>
                </div>

                <div class="pp-step">
                    <span class="pp-step-num">3</span>
                    <strong>Enable Automatic Sending (Optional)</strong>
                    <p style="margin:6px 0 0 32px;">Open your <strong>MorningBatch</strong> script in <strong>Special Content &gt; Python</strong> and add these two lines:</p>
                    <pre style="margin:6px 0 0 32px;background:#f4f4f4;padding:8px 12px;border-radius:4px;font-size:12px;overflow-x:auto;">Data.run_batch = "true"
model.CallScript("TPxi_ProgramPulse")</pre>
                    <p style="margin:6px 0 0 32px;">This runs daily as part of the morning batch. Reports will only send on each configuration's scheduled day.</p>
                </div>

                <h3>Estimated Setup Time</h3>
                <p><strong>5&ndash;10 minutes</strong> for the installation steps above. Another 5 minutes to create your first report configuration. Most of the time is spent deciding which programs/divisions to monitor and who should receive the reports.</p>

                <h3>Content Storage</h3>
                <p>Program Pulse stores its data in three TouchPoint Special Content text entries that are created automatically:</p>
                <ul>
                    <li><code>ProgramPulse_Configs</code> &mdash; All report configurations (programs, recipients, sections, schedules).</li>
                    <li><code>ProgramPulse_Settings</code> &mdash; Global settings (default sender, contact methods).</li>
                    <li><code>ProgramPulse_Log</code> &mdash; Send history (last 100 entries).</li>
                </ul>
                <p style="font-size:12px;color:#888;">These persist independently of the script code. Updating the script will not lose your configurations.</p>

                <h3>Tips</h3>
                <ul>
                    <li>You can create multiple configurations &mdash; for example, one per ministry area, each with different recipients and sections.</li>
                    <li>Use <strong>Hide empty sections</strong> to keep reports clean when there is nothing to report.</li>
                    <li>The <strong>Exclude Organization IDs</strong> field lets you filter out specific organizations you do not want tracked.</li>
                    <li>Check <strong>Send History</strong> to verify reports are going out and to spot errors.</li>
                </ul>
            </div>
        </div>
    </div>

</div>

<!-- EDITOR MODAL -->
<div class="pp-overlay" id="ppEditorOverlay">
    <div class="pp-modal">
        <div class="pp-modal-header">
            <h4 style="margin:0;" id="ppEditorTitle">New Configuration</h4>
            <div style="display:flex;align-items:center;gap:12px;">
                <span class="pp-help-toggle" id="ppHelpAllToggle" onclick="ppToggleAllHelp()" style="width:22px;height:22px;line-height:22px;font-size:13px;" title="Toggle field descriptions">?</span>
                <span style="cursor:pointer;font-size:22px;color:#888;" onclick="ppCloseEditor()">&times;</span>
            </div>
        </div>
        <div class="pp-modal-body">

            <div class="pp-form-group">
                <label>Report Name</label>
                <div class="pp-help-text" id="helpName">A descriptive name for this report configuration. This appears in the email subject line and report header. Example: "Adults Weekly Summary" or "Student Ministry Pulse".</div>
                <input class="pp-input" id="ppCfgName" placeholder="e.g. Adults Weekly Summary">
            </div>

            <div class="pp-row">
                <div class="pp-col-half">
                    <div class="pp-form-group">
                        <label>Lookback Days</label>
                        <div class="pp-help-text" id="helpLookback">How many days back the report should look for activity. For a weekly report, use 7. For bi-weekly use 14. This controls the date range for enrollments, drops, transactions, baptisms, and new people.</div>
                        <input class="pp-input" id="ppCfgLookback" type="number" value="7" min="1" max="90">
                    </div>
                </div>
                <div class="pp-col-half">
                    <div class="pp-form-group">
                        <label>Stale Threshold (days)</label>
                        <div class="pp-help-text" id="helpStale">An organization is considered "stale" if it has had no meeting activity for this many days. Default is 30. Stale involvements may need attention &mdash; they could be inactive groups that were never closed, or groups that stopped meeting.</div>
                        <input class="pp-input" id="ppCfgStale" type="number" value="30" min="7" max="180">
                    </div>
                </div>
            </div>

            <div class="pp-row">
                <div class="pp-col-half">
                    <div class="pp-form-group">
                        <label>Attendance Gap (weeks)</label>
                        <div class="pp-help-text" id="helpGap">The number of recent weeks with zero attendance that qualifies someone as "lapsed." People must have attended 3+ times in the prior 6 months to appear. For example, setting this to 4 means someone who attended regularly but hasn't shown up in 4 weeks will be flagged.</div>
                        <input class="pp-input" id="ppCfgGapWeeks" type="number" value="4" min="1" max="26">
                    </div>
                </div>
                <div class="pp-col-half">
                    <div class="pp-form-group">
                        <label>Scheduled Send Day</label>
                        <div class="pp-help-text" id="helpDay">The day of the week the report will automatically send (via morning batch). The report must also be "Enabled" below. You can always send manually from the Configurations tab regardless of this setting.</div>
                        <div class="pp-day-selector" id="ppCfgDays">
                            <div class="pp-day-btn" data-day="0" onclick="ppSelectDay(this)">Mon</div>
                            <div class="pp-day-btn" data-day="1" onclick="ppSelectDay(this)">Tue</div>
                            <div class="pp-day-btn" data-day="2" onclick="ppSelectDay(this)">Wed</div>
                            <div class="pp-day-btn" data-day="3" onclick="ppSelectDay(this)">Thu</div>
                            <div class="pp-day-btn" data-day="4" onclick="ppSelectDay(this)">Fri</div>
                            <div class="pp-day-btn" data-day="5" onclick="ppSelectDay(this)">Sat</div>
                            <div class="pp-day-btn" data-day="6" onclick="ppSelectDay(this)">Sun</div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="pp-form-group">
                <label>Enabled for Scheduled Sending</label>
                <div class="pp-help-text" id="helpEnabled">When enabled, the report will be automatically sent on the scheduled day via the morning batch. If disabled, the configuration is saved but won't send automatically &mdash; you can still send it manually.</div>
                <label style="font-weight:normal;display:flex;align-items:center;gap:6px;">
                    <input type="checkbox" id="ppCfgEnabled"> Enable automatic sending on scheduled day
                </label>
            </div>

            <div class="pp-form-group">
                <label>Empty Sections</label>
                <div class="pp-help-text" id="helpHideEmpty">When checked, report sections that have no data for the period (e.g., no new baptisms) will be completely hidden from the email. This keeps reports concise. When unchecked, empty sections will still appear with a "nothing to report" message.</div>
                <label style="font-weight:normal;display:flex;align-items:center;gap:6px;">
                    <input type="checkbox" id="ppCfgHideEmpty"> Hide sections with nothing to report
                </label>
            </div>

            <hr style="margin:16px 0;">

            <!-- Program/Division Selection -->
            <div class="pp-form-group">
                <label>Programs &amp; Divisions to Monitor</label>
                <div class="pp-help-text" id="helpProgDiv">Select the programs and divisions whose organizations you want to monitor. You can add a whole program (all divisions) or narrow it to specific divisions. Multiple selections can be combined. Only active organizations within these selections will be included in the report.</div>
                <div class="pp-row">
                    <div class="pp-col-half">
                        <select class="pp-select" id="ppProgSelect">
                            <option value="">-- Select Program --</option>
                        </select>
                    </div>
                    <div class="pp-col-half">
                        <select class="pp-select" id="ppDivSelect" disabled>
                            <option value="">-- Select Division --</option>
                        </select>
                    </div>
                </div>
                <div style="margin-top:8px;">
                    <button class="pp-btn pp-btn-outline pp-btn-sm" onclick="ppAddProgDiv()">+ Add Selection</button>
                </div>
                <div id="ppProgDivTags" style="margin-top:8px;"></div>
            </div>

            <div class="pp-form-group">
                <label>Exclude Organization IDs</label>
                <div class="pp-help-text" id="helpExclude">Comma-separated list of Organization IDs to exclude from monitoring. Useful for filtering out test organizations or special-purpose groups that would add noise to the report. You can find an organization's ID in its URL (e.g., /Organization/<strong>1234</strong>).</div>
                <input class="pp-input" id="ppCfgExclude" placeholder="e.g. 101, 102">
            </div>

            <hr style="margin:16px 0;">

            <!-- Report Sections -->
            <div class="pp-form-group">
                <label>Report Sections</label>
                <div class="pp-help-text" id="helpSections">Choose the detail level for each section. <strong>Detail</strong> shows full tables with individual records and links. <strong>Summary</strong> shows counts and per-organization breakdown. <strong>None</strong> excludes the section entirely. Tailor each report to its audience &mdash; leaders may want detail, executives may prefer summaries.</div>
                <div class="pp-checkbox-group">
                    <label>New Enrollments <select id="ppSec_newEnrollments" class="pp-select" style="width:100px;display:inline-block;margin-left:6px;"><option value="none">None</option><option value="summary">Summary</option><option value="detail" selected>Detail</option></select></label>
                    <label>Dropped Members <select id="ppSec_droppedMembers" class="pp-select" style="width:100px;display:inline-block;margin-left:6px;"><option value="none">None</option><option value="summary">Summary</option><option value="detail" selected>Detail</option></select></label>
                    <label>New Involvements <select id="ppSec_newInvolvements" class="pp-select" style="width:100px;display:inline-block;margin-left:6px;"><option value="none">None</option><option value="summary">Summary</option><option value="detail" selected>Detail</option></select></label>
                    <label>Stale Involvements <select id="ppSec_staleInvolvements" class="pp-select" style="width:100px;display:inline-block;margin-left:6px;"><option value="none">None</option><option value="summary">Summary</option><option value="detail" selected>Detail</option></select></label>
                    <label>Transactions <select id="ppSec_transactions" class="pp-select" style="width:100px;display:inline-block;margin-left:6px;"><option value="none">None</option><option value="summary">Summary</option><option value="detail" selected>Detail</option></select></label>
                    <label>New People <select id="ppSec_newPeople" class="pp-select" style="width:100px;display:inline-block;margin-left:6px;"><option value="none">None</option><option value="summary">Summary</option><option value="detail" selected>Detail</option></select></label>
                    <label>New Baptisms <select id="ppSec_newBaptisms" class="pp-select" style="width:100px;display:inline-block;margin-left:6px;"><option value="none">None</option><option value="summary">Summary</option><option value="detail" selected>Detail</option></select></label>
                    <label>New Church Members <select id="ppSec_newChurchMembers" class="pp-select" style="width:100px;display:inline-block;margin-left:6px;"><option value="none">None</option><option value="summary">Summary</option><option value="detail" selected>Detail</option></select></label>
                    <label>Attendance Gaps <select id="ppSec_attendanceGaps" class="pp-select" style="width:100px;display:inline-block;margin-left:6px;"><option value="none">None</option><option value="summary">Summary</option><option value="detail" selected>Detail</option></select></label>
                </div>
            </div>

            <hr style="margin:16px 0;">

            <!-- Recipients -->
            <div class="pp-form-group">
                <label>Recipients</label>
                <div class="pp-help-text" id="helpRecipients">Search for and add the people who should receive this report by email. Each recipient must have an email address in TouchPoint. You can add multiple recipients &mdash; they each get their own copy of the report.</div>
                <div style="position:relative;">
                    <input class="pp-input" id="ppRecipSearch" placeholder="Search by name or email..." oninput="ppSearchPeople(this.value)">
                    <div class="pp-search-results" id="ppRecipResults" style="display:none;"></div>
                </div>
                <div id="ppRecipTags" style="margin-top:8px;"></div>
            </div>

        </div>
        <div class="pp-modal-footer">
            <button class="pp-btn pp-btn-outline" onclick="ppCloseEditor()">Cancel</button>
            <button class="pp-btn pp-btn-success" onclick="ppSaveConfig()">Save Configuration</button>
        </div>
    </div>
</div>

<script>
var ppConfigs = [];
var ppEditingConfig = null;
var ppProgDivGroups = [];
var ppRecipients = [];
var ppPrograms = [];
var ppDivisions = [];
var ppSearchTimeout = null;

var ppScriptPath = (function() {
    var p = window.location.pathname;
    var parts = p.split('/');
    for (var i = 0; i < parts.length; i++) {
        if (parts[i] === 'PyScriptForm' || parts[i] === 'PyScript') {
            if (i + 1 < parts.length) {
                return '/PyScriptForm/' + parts[i + 1].split('?')[0];
            }
        }
    }
    return p;
})();

function ppAjax(data, callback) {
    $.ajax({
        url: ppScriptPath,
        type: 'POST',
        data: data,
        dataType: 'text',
        success: function(resp) {
            try {
                // Handle case where response may be wrapped in HTML
                var text = resp;
                var jsonStart = text.indexOf('{');
                var jsonEnd = text.lastIndexOf('}');
                if (jsonStart >= 0 && jsonEnd > jsonStart) {
                    text = text.substring(jsonStart, jsonEnd + 1);
                }
                var d = JSON.parse(text);
                callback(d);
            } catch(e) {
                console.error('PP Parse error:', e, resp);
                callback({success: false, message: 'Failed to parse response'});
            }
        },
        error: function(xhr, status, err) {
            console.error('PP AJAX error:', err);
            callback({success: false, message: 'Request failed: ' + err});
        }
    });
}

// ---- INLINE HELP TOGGLE ----
var ppHelpVisible = false;
function ppToggleAllHelp() {
    ppHelpVisible = !ppHelpVisible;
    var helpTexts = document.querySelectorAll('.pp-modal-body .pp-help-text');
    for (var i = 0; i < helpTexts.length; i++) {
        helpTexts[i].className = ppHelpVisible ? 'pp-help-text active' : 'pp-help-text';
    }
    var btn = document.getElementById('ppHelpAllToggle');
    if (btn) {
        btn.style.background = ppHelpVisible ? '#3498db' : '#e8f4fd';
        btn.style.color = ppHelpVisible ? '#fff' : '#3498db';
    }
}

// ---- SENDER BANNER CHECK ----
function ppCheckSenderBanner() {
    ppAjax({action: 'load_settings'}, function(data) {
        var banner = document.getElementById('ppSenderBanner');
        if (!banner) return;
        var hasSender = data.success && data.settings && data.settings.sender && data.settings.sender.peopleId;
        banner.style.display = hasSender ? 'none' : 'flex';
    });
}

// ---- TAB SWITCHING ----
function ppSwitchTab(tab) {
    var tabs = document.querySelectorAll('.pp-tab');
    var contents = document.querySelectorAll('.pp-tab-content');
    for (var i = 0; i < tabs.length; i++) tabs[i].className = 'pp-tab';
    for (var i = 0; i < contents.length; i++) contents[i].className = 'pp-tab-content';
    document.getElementById('ppTab_' + tab).className = 'pp-tab-content active';
    var tabIdx = {configs: 0, preview: 1, log: 2, settings: 3, help: 4};
    tabs[tabIdx[tab] || 0].className = 'pp-tab active';

    if (tab === 'preview') ppPopulatePreviewSelect();
    if (tab === 'log') ppLoadLog();
    if (tab === 'settings') ppLoadSettings();
}

// ---- CONFIG LIST ----
function ppLoadConfigs() {
    ppAjax({action: 'load_configs'}, function(data) {
        if (data.success) {
            ppConfigs = data.configs || [];
            ppConfigs.sort(function(a, b) {
                return (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase());
            });
            ppRenderConfigList();
        }
    });
}

function ppRenderConfigList() {
    var el = document.getElementById('ppConfigList');
    if (!ppConfigs.length) {
        el.innerHTML = '<div class="pp-empty">No configurations yet. Click "+ New Configuration" to create one.</div>';
        return;
    }
    var html = '';
    var days = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
    for (var i = 0; i < ppConfigs.length; i++) {
        var c = ppConfigs[i];
        var progNames = [];
        var groups = c.programDivGroups || [];
        for (var j = 0; j < groups.length; j++) {
            var g = groups[j];
            progNames.push(g.programName + (g.divisionName ? ' > ' + g.divisionName : ' (all divisions)'));
        }
        var dayLabel = c.scheduledDay >= 0 && c.scheduledDay <= 6 ? days[c.scheduledDay] : 'Not set';
        var recipCount = (c.recipients || []).length;
        html += '<div class="pp-config-row">';
        html += '<div>';
        html += '<div class="pp-config-name">' + (c.name || 'Untitled') + '</div>';
        html += '<div class="pp-config-meta">' + progNames.join(', ') + '</div>';
        html += '<div class="pp-config-meta">Schedule: ' + dayLabel + ' &bull; ' + recipCount + ' recipient(s) &bull; ';
        html += '<span class="pp-badge ' + (c.enabled ? 'pp-badge-on' : 'pp-badge-off') + '">' + (c.enabled ? 'Enabled' : 'Disabled') + '</span>';
        html += '</div></div>';
        html += '<div class="pp-config-actions">';
        html += '<button class="pp-btn pp-btn-outline pp-btn-sm" data-id="' + c.id + '" onclick="ppOpenEditor(this.dataset.id)">Edit</button>';
        html += '<button class="pp-btn pp-btn-success pp-btn-sm" data-id="' + c.id + '" onclick="ppSendNow(this.dataset.id)">Send Now</button>';
        html += '<button class="pp-btn pp-btn-danger pp-btn-sm" data-id="' + c.id + '" onclick="ppDeleteConfig(this.dataset.id)">Delete</button>';
        html += '</div></div>';
    }
    el.innerHTML = html;
}

function ppDeleteConfig(id) {
    if (!confirm('Delete this configuration?')) return;
    ppAjax({action: 'delete_config', config_id: id}, function(data) {
        if (data.success) ppLoadConfigs();
        else alert('Error: ' + data.message);
    });
}

function ppSendNow(id) {
    if (!confirm('Send this report now to all recipients?')) return;
    ppAjax({action: 'send_report', config_id: id}, function(data) {
        if (data.success) {
            alert('Report sent to ' + data.sent + ' recipient(s).' + (data.errors.length ? '\\nErrors: ' + data.errors.join(', ') : ''));
        } else {
            alert('Error: ' + data.message);
        }
    });
}

// ---- EDITOR ----
function ppOpenEditor(configId) {
    ppEditingConfig = null;
    ppProgDivGroups = [];
    ppRecipients = [];

    if (configId) {
        for (var i = 0; i < ppConfigs.length; i++) {
            if (ppConfigs[i].id === configId) {
                ppEditingConfig = JSON.parse(JSON.stringify(ppConfigs[i]));
                break;
            }
        }
    }

    if (ppEditingConfig) {
        document.getElementById('ppEditorTitle').textContent = 'Edit Configuration';
        document.getElementById('ppCfgName').value = ppEditingConfig.name || '';
        document.getElementById('ppCfgLookback').value = ppEditingConfig.lookbackDays || 7;
        document.getElementById('ppCfgStale').value = ppEditingConfig.staleThresholdDays || 30;
        document.getElementById('ppCfgGapWeeks').value = ppEditingConfig.attendanceGapWeeks || 4;
        document.getElementById('ppCfgEnabled').checked = !!ppEditingConfig.enabled;
        document.getElementById('ppCfgHideEmpty').checked = !!ppEditingConfig.hideEmptySections;
        document.getElementById('ppCfgExclude').value = ppEditingConfig.excludeOrgIds || '';
        ppProgDivGroups = (ppEditingConfig.programDivGroups || []).slice();
        ppRecipients = (ppEditingConfig.recipients || []).slice();

        // Set scheduled day
        var dayBtns = document.querySelectorAll('#ppCfgDays .pp-day-btn');
        for (var i = 0; i < dayBtns.length; i++) {
            dayBtns[i].className = 'pp-day-btn' + (parseInt(dayBtns[i].getAttribute('data-day')) === ppEditingConfig.scheduledDay ? ' selected' : '');
        }

        // Set sections
        var secs = ppEditingConfig.sections || {};
        var secKeys = ['newEnrollments','droppedMembers','newInvolvements','staleInvolvements','transactions','newPeople','newBaptisms','newChurchMembers','attendanceGaps'];
        for (var i = 0; i < secKeys.length; i++) {
            var sel = document.getElementById('ppSec_' + secKeys[i]);
            if (sel) {
                var val = secs[secKeys[i]];
                // Backwards compatible: true -> detail, false -> none
                if (val === true) val = 'detail';
                else if (val === false) val = 'none';
                else if (!val) val = 'detail';
                sel.value = val;
            }
        }
    } else {
        document.getElementById('ppEditorTitle').textContent = 'New Configuration';
        document.getElementById('ppCfgName').value = '';
        document.getElementById('ppCfgLookback').value = 7;
        document.getElementById('ppCfgStale').value = 30;
        document.getElementById('ppCfgGapWeeks').value = 4;
        document.getElementById('ppCfgEnabled').checked = false;
        document.getElementById('ppCfgHideEmpty').checked = false;
        document.getElementById('ppCfgExclude').value = '';
        var dayBtns = document.querySelectorAll('#ppCfgDays .pp-day-btn');
        for (var i = 0; i < dayBtns.length; i++) dayBtns[i].className = 'pp-day-btn';
        var secSelects = document.querySelectorAll('select[id^="ppSec_"]');
        for (var i = 0; i < secSelects.length; i++) secSelects[i].value = 'detail';
    }

    ppRenderProgDivTags();
    ppRenderRecipTags();
    ppLoadPrograms();
    document.getElementById('ppEditorOverlay').className = 'pp-overlay active';
}

function ppCloseEditor() {
    document.getElementById('ppEditorOverlay').className = 'pp-overlay';
}

function ppSelectDay(el) {
    var btns = document.querySelectorAll('#ppCfgDays .pp-day-btn');
    for (var i = 0; i < btns.length; i++) btns[i].className = 'pp-day-btn';
    el.className = 'pp-day-btn selected';
}

function ppGetSelectedDay() {
    var sel = document.querySelector('#ppCfgDays .pp-day-btn.selected');
    return sel ? parseInt(sel.getAttribute('data-day')) : -1;
}

// ---- PROGRAM/DIVISION SELECTION ----
function ppLoadPrograms() {
    ppAjax({action: 'get_programs'}, function(data) {
        if (data.success) {
            ppPrograms = data.programs;
            var sel = document.getElementById('ppProgSelect');
            sel.innerHTML = '<option value="">-- Select Program --</option>';
            for (var i = 0; i < ppPrograms.length; i++) {
                var p = ppPrograms[i];
                sel.innerHTML += '<option value="' + p.id + '">' + p.name + ' (' + p.orgCount + ' orgs)</option>';
            }
            sel.onchange = function() { ppLoadDivisions(this.value); };
        }
    });
}

function ppLoadDivisions(progId) {
    var divSel = document.getElementById('ppDivSelect');
    if (!progId) {
        divSel.innerHTML = '<option value="">-- Select Division --</option>';
        divSel.disabled = true;
        return;
    }
    ppAjax({action: 'get_divisions', program_id: progId}, function(data) {
        if (data.success) {
            ppDivisions = data.divisions;
            divSel.innerHTML = '<option value="">All Divisions</option>';
            for (var i = 0; i < ppDivisions.length; i++) {
                var d = ppDivisions[i];
                divSel.innerHTML += '<option value="' + d.id + '">' + d.name + ' (' + d.orgCount + ' orgs)</option>';
            }
            divSel.disabled = false;
        }
    });
}

function ppAddProgDiv() {
    var progSel = document.getElementById('ppProgSelect');
    var divSel = document.getElementById('ppDivSelect');
    var progId = progSel.value;
    if (!progId) { alert('Please select a program first.'); return; }
    var progName = progSel.options[progSel.selectedIndex].text.split(' (')[0];
    var divId = divSel.value || null;
    var divName = divId ? divSel.options[divSel.selectedIndex].text.split(' (')[0] : null;

    // Check for duplicates
    for (var i = 0; i < ppProgDivGroups.length; i++) {
        var g = ppProgDivGroups[i];
        if (g.programId == progId && g.divisionId == divId) {
            alert('This selection is already added.');
            return;
        }
    }

    ppProgDivGroups.push({
        programId: parseInt(progId),
        divisionId: divId ? parseInt(divId) : null,
        programName: progName,
        divisionName: divName
    });
    ppRenderProgDivTags();
}

function ppRemoveProgDiv(idx) {
    ppProgDivGroups.splice(idx, 1);
    ppRenderProgDivTags();
}

function ppRenderProgDivTags() {
    var el = document.getElementById('ppProgDivTags');
    if (!ppProgDivGroups.length) {
        el.innerHTML = '<span style="color:#999;font-size:12px;">No programs selected</span>';
        return;
    }
    var html = '';
    for (var i = 0; i < ppProgDivGroups.length; i++) {
        var g = ppProgDivGroups[i];
        var label = g.programName + (g.divisionName ? ' > ' + g.divisionName : ' (all divisions)');
        html += '<span class="pp-tag">' + label + '<span class="pp-tag-remove" onclick="ppRemoveProgDiv(' + i + ')">&times;</span></span>';
    }
    el.innerHTML = html;
}

// ---- RECIPIENT SEARCH ----
function ppSearchPeople(term) {
    if (ppSearchTimeout) clearTimeout(ppSearchTimeout);
    var resultsEl = document.getElementById('ppRecipResults');
    if (term.length < 2) {
        resultsEl.style.display = 'none';
        return;
    }
    ppSearchTimeout = setTimeout(function() {
        ppAjax({action: 'search_people', search_term: term}, function(data) {
            if (data.success && data.people.length) {
                var html = '';
                window._ppSearchResults = data.people;
                for (var i = 0; i < data.people.length; i++) {
                    var p = data.people[i];
                    html += '<div class="pp-search-item" data-idx="' + i + '" onclick="ppClickSearchResult(this)">';
                    html += '<strong>' + p.name + '</strong>';
                    if (p.email) html += ' <span style="color:#888;">(' + p.email + ')</span>';
                    html += '</div>';
                }
                resultsEl.innerHTML = html;
                resultsEl.style.display = 'block';
            } else {
                resultsEl.innerHTML = '<div class="pp-search-item" style="color:#999;">No results found</div>';
                resultsEl.style.display = 'block';
            }
        });
    }, 300);
}

function ppClickSearchResult(el) {
    var sr = window._ppSearchResults[parseInt(el.dataset.idx)];
    ppAddRecipient(sr.peopleId, sr.name, sr.email || '');
}

function ppAddRecipient(pid, name, email) {
    for (var i = 0; i < ppRecipients.length; i++) {
        if (ppRecipients[i].peopleId === pid) return;
    }
    ppRecipients.push({peopleId: pid, name: name, email: email});
    ppRenderRecipTags();
    document.getElementById('ppRecipSearch').value = '';
    document.getElementById('ppRecipResults').style.display = 'none';
}

function ppRemoveRecipient(idx) {
    ppRecipients.splice(idx, 1);
    ppRenderRecipTags();
}

function ppRenderRecipTags() {
    var el = document.getElementById('ppRecipTags');
    if (!ppRecipients.length) {
        el.innerHTML = '<span style="color:#999;font-size:12px;">No recipients added</span>';
        return;
    }
    var html = '';
    for (var i = 0; i < ppRecipients.length; i++) {
        var r = ppRecipients[i];
        html += '<span class="pp-tag">' + r.name;
        if (r.email) html += ' (' + r.email + ')';
        html += '<span class="pp-tag-remove" onclick="ppRemoveRecipient(' + i + ')">&times;</span></span>';
    }
    el.innerHTML = html;
}

// ---- SAVE CONFIG ----
function ppSaveConfig() {
    var name = document.getElementById('ppCfgName').value.trim();
    if (!name) { alert('Please enter a report name.'); return; }
    if (!ppProgDivGroups.length) { alert('Please add at least one program/division.'); return; }

    var secKeys = ['newEnrollments','droppedMembers','newInvolvements','staleInvolvements','transactions','newPeople','newBaptisms','newChurchMembers','attendanceGaps'];
    var sections = {};
    for (var i = 0; i < secKeys.length; i++) {
        var sel = document.getElementById('ppSec_' + secKeys[i]);
        sections[secKeys[i]] = sel ? sel.value : 'detail';
    }

    var config = {
        id: ppEditingConfig ? ppEditingConfig.id : '',
        name: name,
        lookbackDays: parseInt(document.getElementById('ppCfgLookback').value) || 7,
        staleThresholdDays: parseInt(document.getElementById('ppCfgStale').value) || 30,
        attendanceGapWeeks: parseInt(document.getElementById('ppCfgGapWeeks').value) || 4,
        scheduledDay: ppGetSelectedDay(),
        enabled: document.getElementById('ppCfgEnabled').checked,
        hideEmptySections: document.getElementById('ppCfgHideEmpty').checked,
        programDivGroups: ppProgDivGroups,
        excludeOrgIds: document.getElementById('ppCfgExclude').value.trim(),
        sections: sections,
        recipients: ppRecipients
    };

    ppAjax({action: 'save_config', config_data: JSON.stringify(config)}, function(data) {
        if (data.success) {
            ppCloseEditor();
            ppLoadConfigs();
        } else {
            alert('Error: ' + data.message);
        }
    });
}

// ---- PREVIEW ----
function ppPopulatePreviewSelect() {
    var sel = document.getElementById('ppPreviewConfig');
    sel.innerHTML = '<option value="">-- Select Configuration --</option>';
    for (var i = 0; i < ppConfigs.length; i++) {
        sel.innerHTML += '<option value="' + i + '">' + ppConfigs[i].name + '</option>';
    }
}

function ppLoadPreview() {
    var sel = document.getElementById('ppPreviewConfig');
    var idx = sel.value;
    if (idx === '') { return; }
    var config = ppConfigs[parseInt(idx)];
    var area = document.getElementById('ppPreviewArea');
    area.innerHTML = '<div style="text-align:center;padding:40px;"><div class="pp-spinner"></div> Generating report preview...</div>';

    ppAjax({action: 'generate_preview', config_data: JSON.stringify(config)}, function(data) {
        if (data.success) {
            area.innerHTML = data.html;
        } else {
            area.innerHTML = '<div style="color:#c00;padding:20px;">Error: ' + data.message + '</div>';
        }
    });
}

// ---- LOG ----
function ppLoadLog() {
    ppAjax({action: 'get_log'}, function(data) {
        var el = document.getElementById('ppLogArea');
        if (!data.success || !data.logs || !data.logs.length) {
            el.innerHTML = '<div class="pp-empty">No send history yet.</div>';
            return;
        }
        var html = '';
        var logs = data.logs.slice().reverse();
        for (var i = 0; i < logs.length; i++) {
            var l = logs[i];
            var ts = l.timestamp ? l.timestamp.replace('T', ' ').substring(0, 16) : '';
            var statusClass = (l.errors && l.errors.length) ? 'color:#e74c3c;' : 'color:#27ae60;';
            html += '<div class="pp-log-entry">';
            html += '<span style="' + statusClass + 'font-weight:600;">' + (l.source === 'batch' ? 'AUTO' : 'MANUAL') + '</span> ';
            html += '<strong>' + (l.configName || '') + '</strong> &bull; ';
            html += ts + ' &bull; ';
            html += 'Sent to ' + (l.sent || 0) + ' recipient(s)';
            if (l.errors && l.errors.length) {
                html += ' &bull; <span style="color:#e74c3c;">' + l.errors.length + ' error(s)</span>';
            }
            html += '</div>';
        }
        el.innerHTML = html;
    });
}

// ---- SETTINGS ----
var ppKeywords = [];
var ppContactMethods = [];
var ppSender = null; // {peopleId, name, email}
var ppSettingsLoaded = false;
var ppSenderSearchTimeout = null;

function ppLoadSettings() {
    if (ppSettingsLoaded) return;
    ppAjax({action: 'get_keywords'}, function(data) {
        if (data.success) ppKeywords = data.keywords || [];
        ppAjax({action: 'load_settings'}, function(sdata) {
            ppSettingsLoaded = true;
            if (sdata.success && sdata.settings) {
                ppContactMethods = sdata.settings.contact_methods || [];
                ppSender = sdata.settings.sender || null;
            }
            ppRenderContactMethods();
            ppRenderSender();
        });
    });
}

function ppRenderSender() {
    var display = document.getElementById('ppSenderDisplay');
    var input = document.getElementById('ppSenderSearch');
    if (ppSender && ppSender.name) {
        display.innerHTML = ppSender.name + (ppSender.email ? ' (' + ppSender.email + ')' : '') + ' <span style="cursor:pointer;color:#e74c3c;margin-left:6px;" onclick="ppClearSender()">&times;</span>';
        input.value = '';
        input.placeholder = 'Change sender...';
    } else {
        display.textContent = '';
        input.placeholder = 'Search by name...';
    }
}

function ppSearchSender(term) {
    if (ppSenderSearchTimeout) clearTimeout(ppSenderSearchTimeout);
    var resultsEl = document.getElementById('ppSenderResults');
    if (term.length < 2) { resultsEl.style.display = 'none'; return; }
    ppSenderSearchTimeout = setTimeout(function() {
        ppAjax({action: 'search_people', search_term: term}, function(data) {
            if (data.success && data.people && data.people.length) {
                var html = '';
                for (var i = 0; i < data.people.length; i++) {
                    var p = data.people[i];
                    html += '<div class="pp-search-item" data-pid="' + p.peopleId + '" data-name="' + p.name + '" data-email="' + (p.email || '') + '" onclick="ppSelectSender(this)">';
                    html += '<strong>' + p.name + '</strong>';
                    if (p.email) html += ' <span style="color:#888;">(' + p.email + ')</span>';
                    html += '</div>';
                }
                resultsEl.innerHTML = html;
                resultsEl.style.display = 'block';
            } else {
                resultsEl.innerHTML = '<div class="pp-search-item" style="color:#999;">No results</div>';
                resultsEl.style.display = 'block';
            }
        });
    }, 300);
}

function ppSelectSender(el) {
    ppSender = {
        peopleId: parseInt(el.dataset.pid),
        name: el.dataset.name,
        email: el.dataset.email
    };
    document.getElementById('ppSenderResults').style.display = 'none';
    ppRenderSender();
}

function ppClearSender() {
    ppSender = null;
    ppRenderSender();
}

function ppRenderContactMethods() {
    var tbody = document.getElementById('ppContactMethodRows');
    if (!ppContactMethods.length) {
        var emptyHtml = '<tr><td colspan="4" style="padding:12px;color:#999;text-align:center;">No contact methods configured. Click "+ Add Contact Method" to get started.</td></tr>';
        emptyHtml += '<tr style="border-bottom:1px solid #eee;background:#f8f9fa;opacity:0.7;">';
        emptyHtml += '<td style="padding:6px 8px;"><span style="display:inline-block;width:50px;text-align:center;font-weight:700;color:#888;">O</span></td>';
        emptyHtml += '<td style="padding:6px 8px;"><span style="color:#888;">Other</span></td>';
        emptyHtml += '<td style="padding:6px 8px;"><span style="color:#999;font-size:12px;font-style:italic;">Catch-all for any task/note not matching the methods above</span></td>';
        emptyHtml += '<td style="padding:6px 8px;"></td>';
        emptyHtml += '</tr>';
        tbody.innerHTML = emptyHtml;
        return;
    }
    var html = '';
    for (var i = 0; i < ppContactMethods.length; i++) {
        var m = ppContactMethods[i];
        html += '<tr style="border-bottom:1px solid #eee;">';
        html += '<td style="padding:6px 8px;"><input class="pp-input" style="width:50px;text-align:center;font-weight:700;" value="' + (m.code || '') + '" onchange="ppContactMethods[' + i + '].code=this.value" maxlength="2"/></td>';
        html += '<td style="padding:6px 8px;"><input class="pp-input" style="width:120px;" value="' + (m.label || '') + '" onchange="ppContactMethods[' + i + '].label=this.value"/></td>';
        html += '<td style="padding:6px 8px;"><select class="pp-select" onchange="ppSetKeyword(' + i + ',this)">';
        html += '<option value="">-- Select Keyword --</option>';
        for (var j = 0; j < ppKeywords.length; j++) {
            var kw = ppKeywords[j];
            var sel = (m.keywordId && parseInt(m.keywordId) === kw.keywordId) ? ' selected' : '';
            html += '<option value="' + kw.keywordId + '"' + sel + '>' + kw.description + '</option>';
        }
        html += '</select></td>';
        html += '<td style="padding:6px 8px;text-align:center;"><button class="pp-btn pp-btn-danger pp-btn-sm" onclick="ppRemoveContactMethod(' + i + ')">Remove</button></td>';
        html += '</tr>';
    }
    // Fixed O=Other row (always present, not editable)
    html += '<tr style="border-bottom:1px solid #eee;background:#f8f9fa;opacity:0.7;">';
    html += '<td style="padding:6px 8px;"><span style="display:inline-block;width:50px;text-align:center;font-weight:700;color:#888;">O</span></td>';
    html += '<td style="padding:6px 8px;"><span style="color:#888;">Other</span></td>';
    html += '<td style="padding:6px 8px;"><span style="color:#999;font-size:12px;font-style:italic;">Catch-all for any task/note not matching the methods above</span></td>';
    html += '<td style="padding:6px 8px;"></td>';
    html += '</tr>';
    tbody.innerHTML = html;
}

function ppSetKeyword(idx, sel) {
    var opt = sel.options[sel.selectedIndex];
    if (opt.value) {
        ppContactMethods[idx].keywordId = parseInt(opt.value);
        ppContactMethods[idx].keyword = opt.text;
    } else {
        ppContactMethods[idx].keywordId = null;
        ppContactMethods[idx].keyword = '';
    }
}

function ppAddContactMethod() {
    ppContactMethods.push({code: '', label: '', keyword: '', keywordId: null});
    ppRenderContactMethods();
}

function ppRemoveContactMethod(idx) {
    ppContactMethods.splice(idx, 1);
    ppRenderContactMethods();
}

function ppSaveSettings() {
    var settings = {contact_methods: ppContactMethods, sender: ppSender};
    ppAjax({action: 'save_settings', settings_data: JSON.stringify(settings)}, function(data) {
        var el = document.getElementById('ppSettingsStatus');
        if (data.success) {
            el.textContent = 'Saved!';
            setTimeout(function() { el.textContent = ''; }, 3000);
            // Update sender banner visibility
            var banner = document.getElementById('ppSenderBanner');
            if (banner) banner.style.display = (ppSender && ppSender.peopleId) ? 'none' : 'flex';
        } else {
            el.style.color = '#e74c3c';
            el.textContent = 'Error: ' + data.message;
        }
    });
}

// ---- INIT ----
function ppInit() {
    ppLoadConfigs();
    ppCheckSenderBanner();
}
document.addEventListener('DOMContentLoaded', function() {
    ppInit();
});
// Also run immediately in case DOM already loaded
if (document.readyState !== 'loading') {
    ppInit();
}
</script>
'''


# ============================================================
# MAIN ENTRY POINT
# ============================================================

is_batch = (hasattr(Data, 'run_batch') and str(Data.run_batch) == 'true') or model.FromMorningBatch
if is_batch:
    handle_batch()
elif model.HttpMethod == "post":
    action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''
    if action:
        handle_ajax()
    else:
        print json.dumps({'success': False, 'message': 'No action specified'})
else:
    model.Header = TITLE
    model.Form = generate_ui()
