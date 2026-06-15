#roles=Edit

# Written By: Ben Swaby
# Email: bswaby@fbchtn.org
# GitHub:  https://github.com/bswaby/Touchpoint
# ---------------------------------------------------------------
# Support: These tools are free because they should be. If they've
#          saved you time, consider DisplayCache — church digital
#          signage that integrates with TouchPoint.
#          https://displaycache.com
# ---------------------------------------------------------------


#####################################################################
# TPxi Membership Analysis Report  v2.1.2
#####################################################################
#
# SETUP (paste and go):
#   1. Admin > Advanced > Special Content > Python
#   2. Create new script: TPxi_MembershipAnalysisReport
#   3. Paste this entire file and save
#   4. Navigate to /PyScriptForm/TPxi_MembershipAnalysisReport
#   5. Configure settings (gear icon), then Run Report
#
# URLs:
#   Report:   /PyScriptForm/TPxi_MembershipAnalysisReport
#   Settings: /PyScriptForm/TPxi_MembershipAnalysisReport?settings=1
#
# CHANGELOG:
#   v2.1.2 - 2026-06-15
#     - DOC: added a "What these rates actually measure" footnote to the
#       Retention legend. The current SQL measures snapshot retention
#       (still a Member today AND N years passed since join), not
#       point-in-time retention (still a Member at the N-year mark).
#       For older cohorts where everyone who was going to leave has
#       already left, 1Y / 3Y / 5Y collapse to the same number. The
#       footnote names the limitation explicitly so staff don't read
#       into identical adjacent percentages.
#   v2.1.1 - 2026-06-15
#     - FIX: retention showing 100% across every cohort. The cohort CTE
#       filtered `WHERE MemberStatusId = MEMBER`, then the retention CASE
#       checked the same MemberStatusId condition -- so by construction
#       every cohort row was "retained." Removed the cohort-side status
#       and IsDeceased filters; cohort is now "everyone with a JoinDate
#       in the year," and the retention CASE measures `CurrentStatus =
#       MEMBER AND IsDeceased = 0` to find who actually stayed. Real
#       attrition now shows in the rates.
#   v2.1 - 2026-06-15
#     - FIX: Retention rate gating. With the cohort cap removed, the 1/3/5
#       year retention % could show misleadingly low values (early joiners
#       had a year, late joiners had a week). Added get_cohort_year_end_sql()
#       so rate columns only populate once the FULL cohort has had time
#       to mature (cohort year-end + N years has passed). Until then the
#       rate AND the raw count both show "-".
#     - FIX: Retention legend text said "fiscal year" even when Calendar
#       Year was selected. Now toggles between "fiscal year" and "calendar
#       year" based on USE_FISCAL_YEAR. Example years (FY2020 vs 2020)
#       also flip with the setting.
#     - FIX: Baptism age-bin drill-down modal showed "Could not parse
#       response." for every cell. The handler printed raw stdout, but
#       the page is served via /PyScriptForm/ which renders model.Form
#       and ignores raw print. Wrapped handle_baptism_drilldown() in
#       the same StringIO -> model.Form pattern the GET report uses.
#       Now the modal populates with the actual people list.
#   v2.0 - 2026-03-16
#     - Routing overhaul: report is the default view, settings via ?settings=1
#     - Campus selector dropdown in report header for inline switching
#     - Campus filtering applied to all data queries
#     - WITH (NOLOCK) added to all People and ChangeLog queries
#     - YoY delta indicators on yearly membership trends table
#     - Status Transitions section (members gained/lost/returned per year)
#     - Lapsed Members section (high-level annual attrition tracking)
#     - Fiscal year support: toggle fiscal vs calendar year with custom start month/day
#     - Removed CSV/Excel export feature
#     - Removed dead code (safe_int, unused caching config)
#     - Stdout capture to render report inside TouchPoint page frame
#   v1.0 - Initial release
#     - Membership growth trends by year
#     - Age demographics evolution
#     - Retention / attrition cohort analysis
#     - Attendance impact (before/after membership)
#     - Family unit analysis
#     - Baptism age-bin breakdown with drilldown
#     - Campus-specific breakdowns (optional)
#     - Connection pathway analysis (optional)
#     - All settings configurable via UI (no code edits)
#
#####################################################################

import json
import sys
from StringIO import StringIO

# ============================================================
# CONFIGURATION STORAGE
# ============================================================
CONTENT_KEY = "MembershipAnalysisReport_Config"
TITLE = "Membership Analysis Report"
CSS_PREFIX = "mar"

# ============================================================
# DEFAULT CONFIG VALUES
# ============================================================
DEFAULTS = {
    'reportTitle': 'Church Membership Analysis Report',
    'yearsToAnalyze': 5,
    'cohortYearsToAnalyze': 10,
    'fiscalYearStartMonth': 10,
    'fiscalYearStartDay': 1,
    'useFiscalYear': True,
    'memberStatusId': 10,
    'previousMemberStatusId': 40,
    'prospectStatusId': 20,
    'guestStatusId': 30,
    'primaryWorshipProgramId': 1124,
    'primaryWorshipProgramName': 'Worship',
    'smallGroupProgramId': 1128,
    'smallGroupProgramName': 'Connect Groups',
    'campusId': 0,
    'showCampusBreakdown': False,
    'showOriginAnalysis': False,
    'showRetentionMetrics': True,
    'showFamilyAnalysis': True,
    'showAttendanceImpact': True,
    'showBaptismAnalysis': True,
    'showStatusTransitions': True,
    'showLapsedMembers': True,
}

# ============================================================
# CONFIG LOAD / SAVE
# ============================================================
def load_config():
    try:
        content = model.TextContent(CONTENT_KEY)
        if content:
            return json.loads(content)
    except:
        pass
    return {}

def save_config(cfg):
    model.WriteContentText(CONTENT_KEY, json.dumps(cfg, indent=2), "")

def get_cfg(key):
    """Get a config value with fallback to defaults"""
    return _active_config.get(key, DEFAULTS.get(key))

# Load once at module level
_active_config = load_config()

# ============================================================
# CONFIG BRIDGE: maps new config keys to old Config class
# ============================================================
class Config:
    """Bridge class - reads from stored config instead of hardcoded values"""
    @property
    def YEARS_TO_ANALYZE(self):
        return int(get_cfg('yearsToAnalyze'))
    @property
    def COHORT_YEARS_TO_ANALYZE(self):
        return int(get_cfg('cohortYearsToAnalyze'))
    @property
    def REPORT_TITLE(self):
        return get_cfg('reportTitle')
    @property
    def FISCAL_YEAR_START_MONTH(self):
        return int(get_cfg('fiscalYearStartMonth'))
    @property
    def FISCAL_YEAR_START_DAY(self):
        return int(get_cfg('fiscalYearStartDay'))
    @property
    def USE_FISCAL_YEAR(self):
        return bool(get_cfg('useFiscalYear'))
    @property
    def MEMBER_STATUS_ID(self):
        return int(get_cfg('memberStatusId'))
    @property
    def PREVIOUS_MEMBER_STATUS_ID(self):
        return int(get_cfg('previousMemberStatusId'))
    @property
    def PROSPECT_STATUS_ID(self):
        return int(get_cfg('prospectStatusId'))
    @property
    def GUEST_STATUS_ID(self):
        return int(get_cfg('guestStatusId'))
    @property
    def AGE_GROUPS(self):
        return [
            ("Children", 0, 12),
            ("Teens", 13, 17),
            ("Young Adults", 18, 29),
            ("Adults", 30, 49),
            ("Mature Adults", 50, 64),
            ("Seniors", 65, 150)
        ]
    @property
    def PRIMARY_WORSHIP_PROGRAM_ID(self):
        return int(get_cfg('primaryWorshipProgramId'))
    @property
    def PRIMARY_WORSHIP_PROGRAM_NAME(self):
        return get_cfg('primaryWorshipProgramName')
    @property
    def SMALL_GROUP_PROGRAM_ID(self):
        return int(get_cfg('smallGroupProgramId'))
    @property
    def SMALL_GROUP_PROGRAM_NAME(self):
        return get_cfg('smallGroupProgramName')
    @property
    def CAMPUS_ID(self):
        return int(get_cfg('campusId'))
    @property
    def SHOW_CAMPUS_BREAKDOWN(self):
        return bool(get_cfg('showCampusBreakdown'))
    @property
    def SHOW_ORIGIN_ANALYSIS(self):
        return bool(get_cfg('showOriginAnalysis'))
    @property
    def SHOW_RETENTION_METRICS(self):
        return bool(get_cfg('showRetentionMetrics'))
    @property
    def SHOW_FAMILY_ANALYSIS(self):
        return bool(get_cfg('showFamilyAnalysis'))
    @property
    def SHOW_ATTENDANCE_IMPACT(self):
        return bool(get_cfg('showAttendanceImpact'))
    @property
    def SHOW_BAPTISM_ANALYSIS(self):
        return bool(get_cfg('showBaptismAnalysis'))
    @property
    def BAPTISM_AGE_BINS(self):
        return [
            ("Preschool (0-4)", 0, 4),
            ("Children (5-9)", 5, 9),
            ("Preteen (10-11)", 10, 11),
            ("Middle School (12-14)", 12, 14),
            ("High School (15-18)", 15, 18),
            ("19-29", 19, 29),
            ("30-39", 30, 39),
            ("40-49", 40, 49),
            ("50-59", 50, 59),
            ("60-69", 60, 69),
            ("70-79", 70, 79),
            ("80+", 80, 150)
        ]
    @property
    def CHART_COLORS(self):
        return [
            "#667eea", "#48bb78", "#ed8936", "#e53e3e",
            "#38b2ac", "#805ad5", "#d69e2e", "#3182ce",
        ]
    @property
    def SHOW_STATUS_TRANSITIONS(self):
        return bool(get_cfg('showStatusTransitions'))
    @property
    def SHOW_LAPSED_MEMBERS(self):
        return bool(get_cfg('showLapsedMembers'))

Config = Config()

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


# ============================================================
# AJAX HANDLERS
# ============================================================
def handle_ajax():
    action = get_form_data('action', '')
    response = {'success': False, 'message': 'Unknown action'}

    try:
        if action == 'load_config':
            cfg = load_config()
            merged = dict(DEFAULTS)
            merged.update(cfg)
            response = {'success': True, 'config': merged}

        elif action == 'save_config':
            config_json = get_form_data('config_data', '{}')
            cfg = json.loads(config_json)
            cfg['_saved'] = True
            save_config(cfg)
            global _active_config
            _active_config = cfg
            response = {'success': True, 'message': 'Configuration saved'}

        elif action == 'get_campuses':
            sql = """
                SELECT c.Id, c.Description
                FROM lookup.Campus c
                WHERE c.Id > 0
                ORDER BY c.Description
            """
            campuses = [{'id': 0, 'name': 'All Campuses'}]
            try:
                for row in q.QuerySql(sql):
                    campuses.append({
                        'id': row.Id,
                        'name': str(row.Description) if row.Description else 'Campus ' + str(row.Id)
                    })
            except:
                pass
            response = {'success': True, 'campuses': campuses}

        elif action == 'get_programs':
            sql = """
                SELECT p.Id, p.Name
                FROM Program p
                ORDER BY p.Name
            """
            programs = []
            try:
                for row in q.QuerySql(sql):
                    programs.append({
                        'id': row.Id,
                        'name': str(row.Name) if row.Name else 'Program ' + str(row.Id)
                    })
            except:
                pass
            response = {'success': True, 'programs': programs}

    except Exception as e:
        response = {'success': False, 'message': str(e)}

    print json.dumps(response)


# ============================================================
# CONFIG UI
# ============================================================
def generate_ui():
    return '''
<style>
.mar-container { font-family: 'Segoe UI', Tahoma, sans-serif; max-width: 900px; margin: 0 auto; }
.mar-card { background: #fff; border: 1px solid #ddd; border-radius: 6px; margin-bottom: 16px; }
.mar-card-header { background: #f8f9fa; padding: 12px 20px; border-bottom: 1px solid #ddd; border-radius: 6px 6px 0 0; font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
.mar-card-body { padding: 16px 20px; }
.mar-btn { display: inline-block; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: 500; text-decoration: none; }
.mar-btn-primary { background: #3498db; color: #fff; }
.mar-btn-primary:hover { background: #2980b9; }
.mar-btn-success { background: #27ae60; color: #fff; }
.mar-btn-success:hover { background: #219a52; }
.mar-btn-outline { background: #fff; color: #333; border: 1px solid #ccc; }
.mar-btn-outline:hover { background: #f0f0f0; }
.mar-form-group { margin-bottom: 14px; }
.mar-form-group label { display: block; font-weight: 600; margin-bottom: 4px; font-size: 13px; color: #555; }
.mar-input, .mar-select { width: 100%; padding: 8px 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; box-sizing: border-box; }
.mar-input:focus, .mar-select:focus { border-color: #3498db; outline: none; box-shadow: 0 0 0 2px rgba(52,152,219,0.15); }
.mar-row { display: flex; gap: 16px; flex-wrap: wrap; }
.mar-col-half { flex: 1; min-width: 280px; }
.mar-col-third { flex: 1; min-width: 180px; }
.mar-section-title { font-weight: 600; font-size: 14px; margin: 20px 0 10px 0; padding-bottom: 6px; border-bottom: 1px solid #eee; color: #333; }
.mar-checkbox-row { display: flex; align-items: center; gap: 8px; padding: 6px 0; }
.mar-checkbox-row label { font-weight: normal; cursor: pointer; margin: 0; }
.mar-status { padding: 8px 12px; border-radius: 4px; font-size: 13px; margin-top: 10px; display: none; }
.mar-status-success { background: #d5f5e3; color: #27ae60; border: 1px solid #27ae60; }
.mar-status-error { background: #fadbd8; color: #e74c3c; border: 1px solid #e74c3c; }
.mar-help { font-size: 11px; color: #999; margin-top: 2px; }
.mar-welcome { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: #fff; border-radius: 6px; padding: 20px 24px; margin-bottom: 16px; }
.mar-welcome h3 { margin: 0 0 8px 0; font-size: 18px; color: #fff; border: none; padding: 0; }
.mar-welcome p { margin: 0 0 6px 0; font-size: 13px; opacity: 0.95; line-height: 1.5; }
.mar-welcome ol { margin: 8px 0 0 0; padding-left: 20px; font-size: 13px; line-height: 1.8; }
.mar-welcome code { background: rgba(255,255,255,0.2); padding: 1px 5px; border-radius: 3px; font-size: 12px; }
</style>

<div class="mar-container" id="marApp">

    <!-- Welcome / Quick Start (hidden once configured) -->
    <div class="mar-welcome" id="marWelcome" style="display:none;">
        <h3>Welcome to Membership Analysis Report</h3>
        <p>This report analyzes membership trends, demographics, retention, attendance, and more.</p>
        <p><strong>Quick Start:</strong></p>
        <ol>
            <li>Select your <strong>Worship Program</strong> and <strong>Small Group Program</strong> below (required for attendance charts)</li>
            <li>Adjust Fiscal Year settings if your church uses a non-calendar fiscal year</li>
            <li>Toggle report sections on/off as needed</li>
            <li>Click <strong>Save Settings</strong>, then <strong>Run Report</strong></li>
        </ol>
        <p style="margin-top:10px;opacity:0.8;">Settings are saved to your TouchPoint instance and persist across sessions. Other settings have sensible defaults and work out of the box.</p>
    </div>

    <div class="mar-card">
        <div class="mar-card-header">
            <span>Report Configuration</span>
            <div style="display:flex;gap:8px;">
                <button class="mar-btn mar-btn-success" onclick="marRunReport()">Run Report</button>
                <button class="mar-btn mar-btn-primary" onclick="marSaveConfig()">Save Settings</button>
            </div>
        </div>
        <div class="mar-card-body">

            <div id="marStatus" class="mar-status"></div>

            <!-- General Settings -->
            <div class="mar-section-title">General Settings</div>

            <div class="mar-form-group">
                <label>Report Title</label>
                <input class="mar-input" id="marReportTitle" placeholder="Church Membership Analysis Report">
            </div>

            <div class="mar-row">
                <div class="mar-col-half">
                    <div class="mar-form-group">
                        <label>Years to Analyze (Main)</label>
                        <input class="mar-input" id="marYearsToAnalyze" type="number" min="1" max="20" value="5">
                        <div class="mar-help">Number of years for main trend analysis</div>
                    </div>
                </div>
                <div class="mar-col-half">
                    <div class="mar-form-group">
                        <label>Years to Analyze (Cohort Retention)</label>
                        <input class="mar-input" id="marCohortYears" type="number" min="1" max="20" value="10">
                        <div class="mar-help">Longer timeframe for retention cohort analysis</div>
                    </div>
                </div>
            </div>

            <!-- Fiscal Year Settings -->
            <div class="mar-section-title">Fiscal Year Settings</div>

            <div class="mar-row">
                <div class="mar-col-third">
                    <div class="mar-form-group">
                        <label>Use Fiscal Year</label>
                        <select class="mar-select" id="marUseFiscalYear">
                            <option value="true">Yes - Fiscal Year</option>
                            <option value="false">No - Calendar Year</option>
                        </select>
                    </div>
                </div>
                <div class="mar-col-third">
                    <div class="mar-form-group">
                        <label>FY Start Month</label>
                        <select class="mar-select" id="marFYStartMonth">
                            <option value="1">January</option>
                            <option value="2">February</option>
                            <option value="3">March</option>
                            <option value="4">April</option>
                            <option value="5">May</option>
                            <option value="6">June</option>
                            <option value="7">July</option>
                            <option value="8">August</option>
                            <option value="9">September</option>
                            <option value="10">October</option>
                            <option value="11">November</option>
                            <option value="12">December</option>
                        </select>
                    </div>
                </div>
                <div class="mar-col-third">
                    <div class="mar-form-group">
                        <label>FY Start Day</label>
                        <input class="mar-input" id="marFYStartDay" type="number" min="1" max="28" value="1">
                    </div>
                </div>
            </div>

            <!-- Campus Selection -->
            <div class="mar-section-title">Campus Selection</div>

            <div class="mar-form-group">
                <label>Campus</label>
                <select class="mar-select" id="marCampus">
                    <option value="0">All Campuses</option>
                </select>
                <div class="mar-help">Select a specific campus or All Campuses for the full report</div>
            </div>

            <!-- Program IDs -->
            <div class="mar-section-title">Program Configuration</div>

            <div class="mar-row">
                <div class="mar-col-half">
                    <div class="mar-form-group">
                        <label>Worship Program</label>
                        <select class="mar-select" id="marWorshipProgram"></select>
                        <div class="mar-help">Program used for worship headcount data</div>
                    </div>
                </div>
                <div class="mar-col-half">
                    <div class="mar-form-group">
                        <label>Worship Program Display Name</label>
                        <input class="mar-input" id="marWorshipName" value="Worship">
                    </div>
                </div>
            </div>

            <div class="mar-row">
                <div class="mar-col-half">
                    <div class="mar-form-group">
                        <label>Small Group / Connect Group Program</label>
                        <select class="mar-select" id="marSmallGroupProgram"></select>
                        <div class="mar-help">Program used for small group attendance tracking</div>
                    </div>
                </div>
                <div class="mar-col-half">
                    <div class="mar-form-group">
                        <label>Small Group Display Name</label>
                        <input class="mar-input" id="marSmallGroupName" value="Connect Groups">
                    </div>
                </div>
            </div>

            <!-- Report Sections -->
            <div class="mar-section-title">Report Sections</div>
            <div class="mar-help" style="margin-bottom:10px;">Toggle which sections appear in the report</div>

            <div class="mar-checkbox-row">
                <input type="checkbox" id="marShowCampusBreakdown">
                <label for="marShowCampusBreakdown">Campus Breakdown (multi-campus analysis)</label>
            </div>
            <div class="mar-checkbox-row">
                <input type="checkbox" id="marShowOriginAnalysis">
                <label for="marShowOriginAnalysis">Origin Analysis (how people found the church)</label>
            </div>
            <div class="mar-checkbox-row">
                <input type="checkbox" id="marShowRetentionMetrics" checked>
                <label for="marShowRetentionMetrics">Retention Metrics (cohort retention rates)</label>
            </div>
            <div class="mar-checkbox-row">
                <input type="checkbox" id="marShowFamilyAnalysis" checked>
                <label for="marShowFamilyAnalysis">Family Analysis (family unit trends)</label>
            </div>
            <div class="mar-checkbox-row">
                <input type="checkbox" id="marShowAttendanceImpact" checked>
                <label for="marShowAttendanceImpact">Attendance Impact (engagement before/after membership)</label>
            </div>
            <div class="mar-checkbox-row">
                <input type="checkbox" id="marShowBaptismAnalysis" checked>
                <label for="marShowBaptismAnalysis">Baptism Analysis (baptism age bin breakdown)</label>
            </div>
            <div class="mar-checkbox-row">
                <input type="checkbox" id="marShowStatusTransitions" checked>
                <label for="marShowStatusTransitions">Status Transitions (members gained/lost per year)</label>
            </div>
            <div class="mar-checkbox-row">
                <input type="checkbox" id="marShowLapsedMembers" checked>
                <label for="marShowLapsedMembers">Lapsed Members (members lost per year)</label>
            </div>

            <!-- Member Status IDs (advanced) -->
            <div class="mar-section-title">Advanced: Member Status IDs</div>
            <div class="mar-help" style="margin-bottom:10px;">These typically do not need to change. Check lookup.MemberStatus if unsure.</div>

            <div class="mar-row">
                <div class="mar-col-third">
                    <div class="mar-form-group">
                        <label>Member Status ID</label>
                        <input class="mar-input" id="marMemberStatusId" type="number" value="10">
                    </div>
                </div>
                <div class="mar-col-third">
                    <div class="mar-form-group">
                        <label>Previous Member ID</label>
                        <input class="mar-input" id="marPrevMemberStatusId" type="number" value="40">
                    </div>
                </div>
                <div class="mar-col-third">
                    <div class="mar-form-group">
                        <label>Prospect Status ID</label>
                        <input class="mar-input" id="marProspectStatusId" type="number" value="20">
                    </div>
                </div>
            </div>

        </div>
    </div>

</div>

<script>
var marScriptName = 'TPxi_MembershipAnalysisReport';

function marAjax(data, callback) {
    data.ajax = 'true';
    $.ajax({
        url: '/PyScriptForm/' + marScriptName,
        type: 'POST',
        data: data,
        success: function(resp) {
            try {
                var parsed = (typeof resp === 'string') ? JSON.parse(resp) : resp;
                callback(parsed);
            } catch(e) {
                callback({success: false, message: 'Parse error: ' + e.message});
            }
        },
        error: function(xhr, status, err) {
            callback({success: false, message: 'Request failed: ' + err});
        }
    });
}

function marShowStatus(msg, isError) {
    var el = document.getElementById('marStatus');
    el.textContent = msg;
    el.className = 'mar-status ' + (isError ? 'mar-status-error' : 'mar-status-success');
    el.style.display = 'block';
    setTimeout(function() { el.style.display = 'none'; }, 4000);
}

function marCollectConfig() {
    return {
        reportTitle: document.getElementById('marReportTitle').value,
        yearsToAnalyze: parseInt(document.getElementById('marYearsToAnalyze').value) || 5,
        cohortYearsToAnalyze: parseInt(document.getElementById('marCohortYears').value) || 10,
        useFiscalYear: document.getElementById('marUseFiscalYear').value === 'true',
        fiscalYearStartMonth: parseInt(document.getElementById('marFYStartMonth').value) || 10,
        fiscalYearStartDay: parseInt(document.getElementById('marFYStartDay').value) || 1,
        campusId: parseInt(document.getElementById('marCampus').value) || 0,
        primaryWorshipProgramId: parseInt(document.getElementById('marWorshipProgram').value) || 0,
        primaryWorshipProgramName: document.getElementById('marWorshipName').value,
        smallGroupProgramId: parseInt(document.getElementById('marSmallGroupProgram').value) || 0,
        smallGroupProgramName: document.getElementById('marSmallGroupName').value,
        showCampusBreakdown: document.getElementById('marShowCampusBreakdown').checked,
        showOriginAnalysis: document.getElementById('marShowOriginAnalysis').checked,
        showRetentionMetrics: document.getElementById('marShowRetentionMetrics').checked,
        showFamilyAnalysis: document.getElementById('marShowFamilyAnalysis').checked,
        showAttendanceImpact: document.getElementById('marShowAttendanceImpact').checked,
        showBaptismAnalysis: document.getElementById('marShowBaptismAnalysis').checked,
        showStatusTransitions: document.getElementById('marShowStatusTransitions').checked,
        showLapsedMembers: document.getElementById('marShowLapsedMembers').checked,
        memberStatusId: parseInt(document.getElementById('marMemberStatusId').value) || 10,
        previousMemberStatusId: parseInt(document.getElementById('marPrevMemberStatusId').value) || 40,
        prospectStatusId: parseInt(document.getElementById('marProspectStatusId').value) || 20,
        guestStatusId: 30
    };
}

function marApplyConfig(cfg) {
    document.getElementById('marReportTitle').value = cfg.reportTitle || '';
    document.getElementById('marYearsToAnalyze').value = cfg.yearsToAnalyze || 5;
    document.getElementById('marCohortYears').value = cfg.cohortYearsToAnalyze || 10;
    document.getElementById('marUseFiscalYear').value = cfg.useFiscalYear ? 'true' : 'false';
    document.getElementById('marFYStartMonth').value = cfg.fiscalYearStartMonth || 10;
    document.getElementById('marFYStartDay').value = cfg.fiscalYearStartDay || 1;
    document.getElementById('marCampus').value = cfg.campusId || 0;
    document.getElementById('marWorshipName').value = cfg.primaryWorshipProgramName || 'Worship';
    document.getElementById('marSmallGroupName').value = cfg.smallGroupProgramName || 'Connect Groups';
    document.getElementById('marShowCampusBreakdown').checked = !!cfg.showCampusBreakdown;
    document.getElementById('marShowOriginAnalysis').checked = !!cfg.showOriginAnalysis;
    document.getElementById('marShowRetentionMetrics').checked = cfg.showRetentionMetrics !== false;
    document.getElementById('marShowFamilyAnalysis').checked = cfg.showFamilyAnalysis !== false;
    document.getElementById('marShowAttendanceImpact').checked = cfg.showAttendanceImpact !== false;
    document.getElementById('marShowBaptismAnalysis').checked = cfg.showBaptismAnalysis !== false;
    document.getElementById('marShowStatusTransitions').checked = cfg.showStatusTransitions !== false;
    document.getElementById('marShowLapsedMembers').checked = cfg.showLapsedMembers !== false;
    document.getElementById('marMemberStatusId').value = cfg.memberStatusId || 10;
    document.getElementById('marPrevMemberStatusId').value = cfg.previousMemberStatusId || 40;
    document.getElementById('marProspectStatusId').value = cfg.prospectStatusId || 20;
    // Set program selects after they load
    window._marPendingWorshipId = cfg.primaryWorshipProgramId || 0;
    window._marPendingSmallGroupId = cfg.smallGroupProgramId || 0;
}

function marSaveConfig() {
    var cfg = marCollectConfig();
    marAjax({action: 'save_config', config_data: JSON.stringify(cfg)}, function(data) {
        if (data.success) {
            marShowStatus('Settings saved successfully!', false);
            // Hide welcome after first save
            var w = document.getElementById('marWelcome');
            if (w) w.style.display = 'none';
        } else {
            marShowStatus('Error saving: ' + data.message, true);
        }
    });
}

function marRunReport() {
    // Save config first, then redirect to the report page (GET)
    var cfg = marCollectConfig();
    marAjax({action: 'save_config', config_data: JSON.stringify(cfg)}, function(data) {
        if (data.success) {
            window.open('/PyScriptForm/' + marScriptName, '_blank');
        } else {
            marShowStatus('Error saving before run: ' + data.message, true);
        }
    });
}

function marLoadCampuses() {
    marAjax({action: 'get_campuses'}, function(data) {
        if (data.success) {
            var sel = document.getElementById('marCampus');
            sel.innerHTML = '';
            for (var i = 0; i < data.campuses.length; i++) {
                var c = data.campuses[i];
                var opt = document.createElement('option');
                opt.value = c.id;
                opt.textContent = c.name;
                sel.appendChild(opt);
            }
            if (window._marPendingCampusId !== undefined) {
                sel.value = window._marPendingCampusId;
            }
        }
    });
}

function marLoadPrograms() {
    marAjax({action: 'get_programs'}, function(data) {
        if (data.success) {
            var worshipSel = document.getElementById('marWorshipProgram');
            var sgSel = document.getElementById('marSmallGroupProgram');
            worshipSel.innerHTML = '<option value="0">-- Select Program --</option>';
            sgSel.innerHTML = '<option value="0">-- Select Program --</option>';
            for (var i = 0; i < data.programs.length; i++) {
                var p = data.programs[i];
                var opt1 = document.createElement('option');
                opt1.value = p.id;
                opt1.textContent = p.name;
                worshipSel.appendChild(opt1);
                var opt2 = document.createElement('option');
                opt2.value = p.id;
                opt2.textContent = p.name;
                sgSel.appendChild(opt2);
            }
            if (window._marPendingWorshipId) worshipSel.value = window._marPendingWorshipId;
            if (window._marPendingSmallGroupId) sgSel.value = window._marPendingSmallGroupId;
        }
    });
}

function marInit() {
    marAjax({action: 'load_config'}, function(data) {
        var isFirstRun = true;
        if (data.success && data.config) {
            window._marPendingCampusId = data.config.campusId || 0;
            marApplyConfig(data.config);
            // If reportTitle has been changed from default, user has configured before
            if (data.config._saved) isFirstRun = false;
        }
        // Show welcome banner on first run
        var w = document.getElementById('marWelcome');
        if (w && isFirstRun) w.style.display = 'block';
        marLoadCampuses();
        marLoadPrograms();
    });
}

document.addEventListener('DOMContentLoaded', function() { marInit(); });
if (document.readyState !== 'loading') { marInit(); }
</script>
'''


def main():
    """Main execution function"""
    model.Header = Config.REPORT_TITLE
    try:
        # Check permissions
        if not check_permissions():
            return
        
        # Test basic database connectivity first
        try:
            test_count = q.QuerySqlScalar("""
                SELECT COUNT(*) FROM People WHERE MemberStatusId = {} AND IsDeceased = 0
            """.format(Config.MEMBER_STATUS_ID))
            print('<script>console.log("Found {} active members");</script>'.format(test_count))
            
            # Test if JoinDate column exists
            join_date_test = q.QuerySqlScalar("""
                SELECT COUNT(*) FROM People 
                WHERE MemberStatusId = {} 
                  AND IsDeceased = 0 
                  AND JoinDate IS NOT NULL
                  AND JoinDate >= DATEADD(year, -10, GETDATE())
            """.format(Config.MEMBER_STATUS_ID))
            print('<script>console.log("Found {} members with join dates in last 10 years");</script>'.format(join_date_test))
            
        except Exception as e:
            print('<div class="alert alert-danger">Database connectivity test failed: {}</div>'.format(str(e)))
            print('<div class="alert alert-info">This report requires access to People table with JoinDate field.</div>')
            return
        
        # Calculate date range
        end_date = model.DateTime
        start_date_str = q.QuerySqlScalar("""
            SELECT CONVERT(varchar, DATEADD(year, -{}, GETDATE()), 120)
        """.format(Config.YEARS_TO_ANALYZE))
        
        # For cohort analysis, use longer timeframe
        cohort_start_date_str = q.QuerySqlScalar("""
            SELECT CONVERT(varchar, DATEADD(year, -{}, GETDATE()), 120)
        """.format(Config.COHORT_YEARS_TO_ANALYZE))
        
        print('<script>console.log("Date range: {} to {}");</script>'.format(start_date_str, end_date.ToString("yyyy-MM-dd")))
        
        # Show loading message
        print('<div id="loadingMessage" class="alert alert-info"><i class="fa fa-spinner fa-spin"></i> Generating {} Year Analysis Report (with {} Year Cohort Analysis)... This may take a moment.</div>'.format(Config.YEARS_TO_ANALYZE, Config.COHORT_YEARS_TO_ANALYZE))
        
        # Gather all analytics data (pass cohort start date as well)
        analytics = gather_analytics_data(start_date_str, end_date.ToString("yyyy-MM-dd"), cohort_start_date_str)
        
        # Generate report
        generate_report(analytics, start_date_str, end_date.ToString("yyyy-MM-dd"))
        
        # Hide loading message
        print('<script>document.getElementById("loadingMessage").style.display = "none";</script>')
        
    except Exception as e:
        print('<div class="alert alert-danger">Error generating report: {}</div>'.format(str(e)))
        if model.UserIsInRole("Developer"):
            import traceback
            print('<pre>{}</pre>'.format(traceback.format_exc()))

def check_permissions():
    """Check if user has required permissions"""
    if not model.UserIsInRole("Edit") and not model.UserIsInRole("Admin"):
        print('<div class="alert alert-danger"><i class="fa fa-lock"></i> You need Edit or Admin role to access this report.</div>')
        return False
    return True

def gather_analytics_data(start_date, end_date, cohort_start_date_str=None):
    """Gather all analytics data for the report"""
    analytics = {}
    
    # Use cohort_start_date_str for retention metrics if provided
    if not cohort_start_date_str:
        cohort_start_date_str = start_date
    
    try:
        # 1. Overall membership trends by year
        print('<script>console.log("Loading yearly trends...");</script>')
        analytics['yearly_trends'] = get_yearly_membership_trends(start_date, end_date)
        
        # 2. Age demographics over time
        print('<script>console.log("Loading age demographics...");</script>')
        analytics['age_demographics'] = get_age_demographics_trends(start_date, end_date)
        
        # 3. Connection pathways
        if Config.SHOW_ORIGIN_ANALYSIS:
            print('<script>console.log("Loading origin trends...");</script>')
            analytics['origins'] = get_origin_trends(start_date, end_date)
        
        # 4. Attendance impact analysis
        if Config.SHOW_ATTENDANCE_IMPACT:
            print('<script>console.log("Loading attendance impact...");</script>')
            analytics['attendance_impact'] = get_attendance_impact_analysis(start_date, end_date)
        
        # 5. Family analysis
        if Config.SHOW_FAMILY_ANALYSIS:
            print('<script>console.log("Loading family trends...");</script>')
            analytics['family_trends'] = get_family_trends(start_date, end_date)
        
        # 6. Retention metrics (use cohort years)
        if Config.SHOW_RETENTION_METRICS:
            print('<script>console.log("Loading retention metrics...");</script>')
            analytics['retention'] = get_retention_metrics(cohort_start_date_str, end_date)
        
        # 7. Campus breakdown
        if Config.SHOW_CAMPUS_BREAKDOWN and Config.CAMPUS_ID == 0:
            print('<script>console.log("Loading campus breakdown...");</script>')
            analytics['campus_breakdown'] = get_campus_breakdown(start_date, end_date)
        
        # 8. Baptism age bin analysis
        if Config.SHOW_BAPTISM_ANALYSIS:
            print('<script>console.log("Loading baptism age analysis...");</script>')
            analytics['baptism_age'] = get_baptism_age_trends(start_date, end_date)

        # 9. Current totals
        print('<script>console.log("Loading current totals...");</script>')
        analytics['current_totals'] = get_current_totals()

        # 10. Lapsed members (high-level counts by year)
        if Config.SHOW_LAPSED_MEMBERS:
            print('<script>console.log("Loading lapsed member stats...");</script>')
            analytics['lapsed'] = get_lapsed_member_stats(start_date, end_date)

        # 11. Member status transitions (high-level flows)
        if Config.SHOW_STATUS_TRANSITIONS:
            print('<script>console.log("Loading status transitions...");</script>')
            analytics['transitions'] = get_status_transitions(start_date, end_date)
        
    except Exception as e:
        print('<div class="alert alert-warning">Error loading analytics data: {}</div>'.format(str(e)))
        raise
    
    return analytics

def get_yearly_membership_trends(start_date, end_date):
    """Get membership trends by year"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    sql = """
    SELECT 
        {} AS JoinYear,
        COUNT(*) AS NewMembers,
        COUNT(DISTINCT p.FamilyId) AS NewFamilies,
        -- Gender breakdown
        SUM(CASE WHEN p.GenderId = 1 THEN 1 ELSE 0 END) AS Males,
        SUM(CASE WHEN p.GenderId = 2 THEN 1 ELSE 0 END) AS Females,
        -- Age at joining (using current age as approximation)
        AVG(p.Age) AS AvgAgeAtJoining,
        -- Marital status
        SUM(CASE WHEN p.MaritalStatusId = 20 THEN 1 ELSE 0 END) AS Married,
        SUM(CASE WHEN p.MaritalStatusId = 10 THEN 1 ELSE 0 END) AS Single
    FROM People p WITH (NOLOCK)
    WHERE p.MemberStatusId = {}
      AND p.JoinDate >= '{}'
      AND p.JoinDate <= '{}'
      AND p.IsDeceased = 0{}
    GROUP BY {}
    ORDER BY {} DESC
    """.format(
        fiscal_year_sql,
        Config.MEMBER_STATUS_ID,
        start_date,
        end_date,
        campus_filter(),
        fiscal_year_sql,
        fiscal_year_sql
    )

    return q.QuerySql(sql)

def get_age_demographics_trends(start_date, end_date):
    """Get age demographics trends over time"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    sql = """
    SELECT 
        {} AS JoinYear,
    """.format(fiscal_year_sql)
    
    # Add age group columns dynamically
    for group_name, min_age, max_age in Config.AGE_GROUPS:
        sql += """
        SUM(CASE WHEN p.Age BETWEEN {} AND {} THEN 1 ELSE 0 END) AS {},
        """.format(min_age, max_age, group_name.replace(" ", ""))
    
    sql += """
        SUM(CASE WHEN p.Age IS NULL THEN 1 ELSE 0 END) AS UnknownAge
    FROM People p WITH (NOLOCK)
    WHERE p.MemberStatusId = {}
      AND p.JoinDate >= '{}'
      AND p.JoinDate <= '{}'
      AND p.IsDeceased = 0{}
    GROUP BY {}
    ORDER BY {} DESC
    """.format(
        Config.MEMBER_STATUS_ID,
        start_date,
        end_date,
        campus_filter(),
        fiscal_year_sql,
        fiscal_year_sql
    )
    
    return q.QuerySql(sql)

def get_origin_trends(start_date, end_date):
    """Get trends in how people found the church"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    # Note: Origin table may not exist in all TouchPoint instances
    # Fallback to simple count by year if Origin data is not available
    try:
        sql = """
        SELECT 
            {} AS JoinYear,
            ISNULL(o.Description, 'Not Specified') AS Origin,
            COUNT(*) AS Count
        FROM People p WITH (NOLOCK)
        LEFT JOIN Origin o ON p.OriginId = o.Id
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0{}
        GROUP BY {}, o.Description
        ORDER BY {} DESC, COUNT(*) DESC
        """.format(
            fiscal_year_sql,
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date,
            campus_filter(),
            fiscal_year_sql,
            fiscal_year_sql
        )
        
        return q.QuerySql(sql)
    except:
        # Fallback if Origin table doesn't exist
        sql = """
        SELECT 
            {} AS JoinYear,
            'Not Available' AS Origin,
            COUNT(*) AS Count
        FROM People p WITH (NOLOCK)
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0{}
        GROUP BY {}
        ORDER BY {}, COUNT(*) DESC
        """.format(
            fiscal_year_sql,
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date,
            campus_filter(),
            fiscal_year_sql,
            fiscal_year_sql
        )
        
        return q.QuerySql(sql)

def get_attendance_impact_analysis(start_date, end_date):
    """Analyze attendance patterns before and after membership
    
    This combines:
    1. Worship headcount data (from Meetings table) - using MaxCount for total attendance
    2. Connect Group individual attendance (from Attend table)
    
    It compares attendance patterns 6 months before and after membership.
    """
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    # First check if we can get worship headcount data
    # Check both HeadCount and MaxCount columns
    try:
        headcount_test = q.QuerySqlScalar("""
            SELECT COUNT(*) 
            FROM Meetings m
            WHERE (m.HeadCount > 0 OR m.MaxCount > 0)
              AND m.MeetingDate >= DATEADD(month, -1, GETDATE())
        """)
        use_headcount = headcount_test > 0
        
        # Debug: Check worship meetings specifically
        worship_meeting_test = q.QuerySqlTop1("""
            SELECT 
                COUNT(*) as TotalMeetings,
                SUM(CASE WHEN m.MaxCount > 0 THEN 1 ELSE 0 END) as WithMaxCount,
                SUM(CASE WHEN m.HeadCount > 0 THEN 1 ELSE 0 END) as WithHeadCount,
                AVG(CASE WHEN m.MaxCount > 0 THEN m.MaxCount ELSE m.HeadCount END) as AvgCount
            FROM Meetings m
            INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
            WHERE os.ProgId = {}
              AND m.MeetingDate >= DATEADD(month, -3, GETDATE())
        """.format(Config.PRIMARY_WORSHIP_PROGRAM_ID))
        
        if worship_meeting_test:
            print('<script>console.log("Worship meetings last 3 months: Total={}, WithMaxCount={}, WithHeadCount={}, AvgCount={}");</script>'.format(
                safe_get_value(worship_meeting_test, 'TotalMeetings', 0),
                safe_get_value(worship_meeting_test, 'WithMaxCount', 0),
                safe_get_value(worship_meeting_test, 'WithHeadCount', 0),
                int(safe_get_value(worship_meeting_test, 'AvgCount', 0))
            ))
            
        # Also check just the meetings table structure
        sample_meeting = q.QuerySqlTop1("""
            SELECT TOP 1 m.*, o.OrganizationName
            FROM Meetings m
            INNER JOIN Organizations o ON m.OrganizationId = o.OrganizationId
            INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
            WHERE os.ProgId = {}
              AND m.MeetingDate >= DATEADD(month, -1, GETDATE())
              AND (m.MaxCount > 0 OR m.HeadCount > 0)
            ORDER BY m.MeetingDate DESC
        """.format(Config.PRIMARY_WORSHIP_PROGRAM_ID))
        
        if sample_meeting:
            print('<script>console.log("Sample worship meeting - Org: {}, MaxCount: {}, HeadCount: {}");</script>'.format(
                safe_get_value(sample_meeting, 'OrganizationName', 'Unknown'),
                safe_get_value(sample_meeting, 'MaxCount', 0),
                safe_get_value(sample_meeting, 'HeadCount', 0)
            ))
    except Exception as e:
        use_headcount = False
        print('<script>console.log("Error checking worship meetings: {}");</script>'.format(str(e).replace('"', '\"')))
    
    # Get worship data separately to avoid aggregate function errors
    worship_data = {}
    if use_headcount:
        # Get fiscal years first
        year_sql = """
        SELECT DISTINCT {} AS Year
        FROM People p WITH (NOLOCK)
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0{}
        ORDER BY Year
        """.format(
            get_fiscal_year_sql().format("p.JoinDate"),
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date,
            campus_filter()
        )
        
        years = q.QuerySql(year_sql)
        
        # For each year, get average worship headcount
        for year_row in years:
            year = safe_get_value(year_row, 'Year', None)
            if not year:
                continue
                
            # Convert year to integer to avoid "invalid integer number literal" errors
            try:
                year_int = int(year)
            except (ValueError, TypeError):
                print('<script>console.log("Invalid year value - skipping");</script>')
                continue
            
            # Calculate fiscal year date range
            # Python 2.7.3 doesn't support {:02d} format
            if Config.USE_FISCAL_YEAR:
                if Config.FISCAL_YEAR_START_MONTH >= 10:  # Oct, Nov, Dec
                    month_str = str(Config.FISCAL_YEAR_START_MONTH).zfill(2)
                    day_str = str(Config.FISCAL_YEAR_START_DAY).zfill(2)
                    end_month_str = str(Config.FISCAL_YEAR_START_MONTH - 1).zfill(2)
                    year_start = "{}-{}-{}".format(year_int - 1, month_str, day_str)
                    year_end = "{}-{}-30".format(year_int, end_month_str)
                else:
                    month_str = str(Config.FISCAL_YEAR_START_MONTH).zfill(2)
                    day_str = str(Config.FISCAL_YEAR_START_DAY).zfill(2)
                    end_month_str = str(Config.FISCAL_YEAR_START_MONTH - 1).zfill(2)
                    year_start = "{}-{}-{}".format(year_int, month_str, day_str)
                    year_end = "{}-{}-28".format(year_int + 1, end_month_str)
            else:
                year_start = "{}-01-01".format(year_int)
                year_end = "{}-12-31".format(year_int)
            
            # Debug: Let's see what organizations are under worship
            debug_sql = """
            SELECT COUNT(*) as OrgCount, COUNT(DISTINCT os.OrgId) as UniqueOrgs
            FROM OrganizationStructure os
            WHERE os.ProgId = {}
            """.format(Config.PRIMARY_WORSHIP_PROGRAM_ID)
            
            try:
                org_count = q.QuerySqlTop1(debug_sql)
                org_count_val = safe_get_value(org_count, 'OrgCount', 0)
                print('<script>console.log("Worship orgs for year " + "{}" + ": " + "{}" + " orgs");</script>'.format(year_int, int(org_count_val) if org_count_val else 0))
            except:
                pass
            
            # Updated worship query - aggregate by week to get typical Sunday attendance
            # Since worship is typically once per week, we sum by week and then average those weekly totals
            worship_sql = """
            WITH WeeklyTotals AS (
                SELECT 
                    DATEPART(year, m.MeetingDate) as MeetingYear,
                    DATEPART(week, m.MeetingDate) as MeetingWeek,
                    SUM(CASE 
                        WHEN m.MaxCount > 0 THEN m.MaxCount 
                        WHEN m.HeadCount > 0 THEN m.HeadCount
                        ELSE 0 
                    END) AS WeeklyTotal
                FROM Meetings m
                INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                WHERE m.MeetingDate >= '{}'
                  AND m.MeetingDate <= '{}'
                  AND os.ProgId = {}
                GROUP BY DATEPART(year, m.MeetingDate), DATEPART(week, m.MeetingDate)
            )
            SELECT AVG(CAST(WeeklyTotal AS FLOAT)) AS AvgHeadcount
            FROM WeeklyTotals
            WHERE WeeklyTotal > 0
            """.format(year_start, year_end, Config.PRIMARY_WORSHIP_PROGRAM_ID)
            
            try:
                avg_headcount = q.QuerySqlScalar(worship_sql)
                if avg_headcount:
                    worship_data[year_int] = float(avg_headcount)
                    avg_int = int(avg_headcount) if avg_headcount else 0
                    print('<script>console.log("Year " + "{}" + " worship avg: " + "{}");</script>'.format(year_int, avg_int))
                    
                    # Also get meeting count for context
                    # Get count of unique weeks with worship meetings
                    week_count_sql = """
                    SELECT COUNT(DISTINCT DATEPART(year, m.MeetingDate) * 100 + DATEPART(week, m.MeetingDate)) as UniqueWeeks
                    FROM Meetings m
                    INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                    WHERE m.MeetingDate >= '{}'
                      AND m.MeetingDate <= '{}'
                      AND os.ProgId = {}
                    """.format(year_start, year_end, Config.PRIMARY_WORSHIP_PROGRAM_ID)
                    week_count = q.QuerySqlScalar(week_count_sql)
                    week_count_val = int(week_count) if week_count else 0
                    print('<script>console.log("Year " + "{}" + " had " + "{}" + " worship weeks");</script>'.format(year_int, week_count_val))
                else:
                    # Try a simpler query to see if there's any data
                    simple_test = q.QuerySqlScalar("""
                        SELECT COUNT(*) 
                        FROM Meetings m
                        INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                        WHERE m.MeetingDate >= '{}'
                          AND m.MeetingDate <= '{}'
                          AND os.ProgId = {}
                    """.format(year_start, year_end, Config.PRIMARY_WORSHIP_PROGRAM_ID))
                    simple_test_val = int(simple_test) if simple_test else 0
                    print('<script>console.log("Year " + "{}" + " has " + "{}" + " worship meetings but no headcount data");</script>'.format(year_int, simple_test_val))
                    
                    # Let's check a sample of what MaxCount/HeadCount values look like
                    sample_sql = """
                    SELECT TOP 5 m.MeetingDate, o.OrganizationName, m.MaxCount, m.HeadCount
                    FROM Meetings m
                    INNER JOIN Organizations o ON m.OrganizationId = o.OrganizationId
                    INNER JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
                    WHERE m.MeetingDate >= '{}'
                      AND m.MeetingDate <= '{}'
                      AND os.ProgId = {}
                    ORDER BY m.MeetingDate DESC
                    """.format(year_start, year_end, Config.PRIMARY_WORSHIP_PROGRAM_ID)
                    
                    samples = q.QuerySql(sample_sql)
                    if samples:
                        for s in samples[:2]:  # Just show first 2
                            meeting_date = safe_get_value(s, 'MeetingDate', 'Unknown')
                            org_name = safe_get_value(s, 'OrganizationName', 'Unknown')
                            max_count = int(safe_get_value(s, 'MaxCount', 0))
                            head_count = int(safe_get_value(s, 'HeadCount', 0))
                            print('<script>console.log("Sample: {} - {} - MaxCount: {}, HeadCount: {}");</script>'.format(
                                str(meeting_date), str(org_name), max_count, head_count))
            except Exception as e:
                error_msg = str(e).replace('"', '').replace("'", '')
                print('<script>console.log("Error getting worship data for year " + "{}" + ": " + "{}");</script>'.format(year_int, error_msg))
                pass
    
    # Now get connect group and overall ministry attendance data
    sql = """
    WITH MembershipData AS (
        SELECT
            p.PeopleId,
            p.JoinDate,
            {} AS JoinYear
        FROM People p WITH (NOLOCK)
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0{}
    ),
    EngagementData AS (
        SELECT 
            md.PeopleId,
            md.JoinYear,
            -- Connect Group attendance 6 months before/after joining
            SUM(CASE WHEN a.MeetingDate >= DATEADD(month, -6, md.JoinDate) 
                     AND a.MeetingDate < md.JoinDate
                     AND os.ProgId = {}
                THEN 1 ELSE 0 END) AS CGAttendanceBefore,
            SUM(CASE WHEN a.MeetingDate >= md.JoinDate 
                     AND a.MeetingDate < DATEADD(month, 6, md.JoinDate)
                     AND os.ProgId = {}
                THEN 1 ELSE 0 END) AS CGAttendanceAfter,
            -- Overall ministry engagement (all programs except worship) 6 months before/after
            SUM(CASE WHEN a.MeetingDate >= DATEADD(month, -6, md.JoinDate) 
                     AND a.MeetingDate < md.JoinDate
                     AND os.ProgId != {}
                THEN 1 ELSE 0 END) AS AllMinistryBefore,
            SUM(CASE WHEN a.MeetingDate >= md.JoinDate 
                     AND a.MeetingDate < DATEADD(month, 6, md.JoinDate)
                     AND os.ProgId != {}
                THEN 1 ELSE 0 END) AS AllMinistryAfter
        FROM MembershipData md
        LEFT JOIN Attend a ON md.PeopleId = a.PeopleId AND a.AttendanceFlag = 1
        LEFT JOIN Meetings m ON a.MeetingId = m.MeetingId
        LEFT JOIN OrganizationStructure os ON os.OrgId = m.OrganizationId
        GROUP BY md.PeopleId, md.JoinYear
    )
    SELECT 
        JoinYear,
        COUNT(DISTINCT PeopleId) AS TotalMembers,
        -- Connect Group metrics
        AVG(CAST(CGAttendanceBefore AS FLOAT)) AS AvgSmallGroupBefore,
        AVG(CAST(CGAttendanceAfter AS FLOAT)) AS AvgSmallGroupAfter,
        COUNT(CASE WHEN CGAttendanceAfter > CGAttendanceBefore THEN 1 END) AS ImprovedSmallGroupCount,
        -- Overall ministry metrics (excluding worship)
        AVG(CAST(AllMinistryBefore AS FLOAT)) AS AvgMinistryBefore,
        AVG(CAST(AllMinistryAfter AS FLOAT)) AS AvgMinistryAfter,
        COUNT(CASE WHEN AllMinistryAfter > AllMinistryBefore THEN 1 END) AS ImprovedMinistryCount
    FROM EngagementData
    GROUP BY JoinYear
    ORDER BY JoinYear DESC
    """.format(
        fiscal_year_sql,
        Config.MEMBER_STATUS_ID,
        start_date,
        end_date,
        campus_filter(),
        Config.SMALL_GROUP_PROGRAM_ID,
        Config.SMALL_GROUP_PROGRAM_ID,
        Config.PRIMARY_WORSHIP_PROGRAM_ID,
        Config.PRIMARY_WORSHIP_PROGRAM_ID
    )
    
    results = q.QuerySql(sql)
    
    # Merge worship data with connect group data
    final_results = []
    for row in results:
        # Create a new object with all the data
        join_year = safe_get_value(row, 'JoinYear', None)
        result = type('obj', (object,), {
            'JoinYear': join_year,
            'TotalMembers': safe_get_value(row, 'TotalMembers', 0),
            'AvgWorshipHeadcount': worship_data.get(join_year, 0),
            'AvgSmallGroupBefore': safe_get_value(row, 'AvgSmallGroupBefore', 0),
            'AvgSmallGroupAfter': safe_get_value(row, 'AvgSmallGroupAfter', 0),
            'ImprovedSmallGroupCount': safe_get_value(row, 'ImprovedSmallGroupCount', 0),
            'AvgMinistryBefore': safe_get_value(row, 'AvgMinistryBefore', 0),
            'AvgMinistryAfter': safe_get_value(row, 'AvgMinistryAfter', 0),
            'ImprovedMinistryCount': safe_get_value(row, 'ImprovedMinistryCount', 0)
        })()
        final_results.append(result)
    
    return final_results

def get_family_trends(start_date, end_date):
    """Analyze family unit trends"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    sql = """
    WITH FamilyData AS (
        SELECT 
            FamilyId,
            COUNT(*) AS FamilySize,
            MAX(CASE WHEN Age < 18 THEN 1 ELSE 0 END) AS HasChildren
        FROM People WITH (NOLOCK)
        WHERE IsDeceased = 0
        GROUP BY FamilyId
    )
    SELECT 
        {} AS JoinYear,
        COUNT(DISTINCT p.FamilyId) AS UniqueFamilies,
        -- Family size categories
        COUNT(DISTINCT CASE WHEN f.FamilySize = 1 THEN p.FamilyId END) AS SinglePersonFamilies,
        COUNT(DISTINCT CASE WHEN f.FamilySize = 2 THEN p.FamilyId END) AS CouplesFamilies,
        COUNT(DISTINCT CASE WHEN f.FamilySize BETWEEN 3 AND 4 THEN p.FamilyId END) AS SmallFamilies,
        COUNT(DISTINCT CASE WHEN f.FamilySize >= 5 THEN p.FamilyId END) AS LargeFamilies,
        -- Family composition
        COUNT(DISTINCT CASE WHEN f.HasChildren = 1 THEN p.FamilyId END) AS FamiliesWithChildren,
        AVG(CAST(f.FamilySize AS FLOAT)) AS AvgFamilySize
    FROM People p WITH (NOLOCK)
    INNER JOIN FamilyData f ON p.FamilyId = f.FamilyId
    WHERE p.MemberStatusId = {}
      AND p.JoinDate >= '{}'
      AND p.JoinDate <= '{}'
      AND p.IsDeceased = 0{}
    GROUP BY {}
    ORDER BY {} DESC
    """.format(
        fiscal_year_sql,
        Config.MEMBER_STATUS_ID,
        start_date,
        end_date,
        campus_filter(),
        fiscal_year_sql,
        fiscal_year_sql
    )
    
    return q.QuerySql(sql)

def get_cohort_year_end_sql(year_expr):
    """
    SQL expression for the END date of a given cohort year.
    Used to gate the 1Y / 3Y / 5Y retention rate columns -- a rate is
    only meaningful once *every* member of the cohort has had at least
    N years to either stay or leave.

    Calendar year N -> ends Dec 31, N
    Fiscal year N (start month M, day D) -> ends day before M/D of year N
       e.g. FY 2025 with Oct 1 start -> Sep 30, 2025
    """
    if not Config.USE_FISCAL_YEAR:
        # Calendar: day before Jan 1 of next year = Dec 31 of cohort year
        return "DATEADD(day, -1, CAST(CAST(({0})+1 AS varchar) + '-01-01' AS datetime))".format(year_expr)
    return ("DATEADD(day, -1, CAST(CAST(({0}) AS varchar) + '-{1:02d}-{2:02d}' AS datetime))"
            .format(year_expr, Config.FISCAL_YEAR_START_MONTH, Config.FISCAL_YEAR_START_DAY))


def get_retention_metrics(start_date, end_date):
    """Calculate retention and attrition metrics.

    Cohort size = everyone who joined in that year, period. The earlier
    `JoinDate <= DATEADD(year, -1, GETDATE())` filter silently dropped
    anyone who joined in the last 12 months from the cohort itself,
    undercounting the most recent cohort (e.g. a 96-member 2025 cohort
    showed as 37 when run on 2026-06-15). Bug found 2026-06-15.

    CRITICAL (added 2026-06-15 round 2): The cohort CTE must NOT filter
    on MemberStatusId or IsDeceased. The earlier version filtered the
    cohort to currently-Members-and-alive, then the retention CASE
    counted "still a Member" -- the same condition on both sides made
    retention 100% by construction across every row. The cohort is
    "everyone who joined in that year," irrespective of where they are
    today. Retention then measures what fraction is STILL a Member and
    still alive.

    Retention 1Y / 3Y / 5Y rates are only shown once the FULL cohort
    has had at least N years to mature -- i.e. (cohort year-end + N)
    is in the past. Otherwise they'd be apples-to-oranges (early
    joiners get a year, late joiners get a week).
    """
    fiscal_year_sql = get_fiscal_year_sql().format("JoinDate")
    cohort_end = get_cohort_year_end_sql("CohortYear")

    sql = """
    WITH MemberCohorts AS (
        -- Everyone with a JoinDate in the year, REGARDLESS of current
        -- MemberStatusId or IsDeceased. We carry both forward so the
        -- retention CASE can measure attrition (left-the-church AND
        -- died-since count as not-retained). Campus filter still applies.
        SELECT
            {0} AS CohortYear,
            PeopleId,
            JoinDate,
            MemberStatusId AS CurrentStatus,
            IsDeceased
        FROM People WITH (NOLOCK)
        WHERE JoinDate IS NOT NULL
          AND JoinDate >= '{2}'
          AND JoinDate <= GETDATE(){3}
    ),
    RetentionData AS (
        SELECT
            mc.CohortYear,
            COUNT(DISTINCT mc.PeopleId) AS CohortSize,
            -- 1 year retention: still a Member, still alive, 1Y anniversary passed
            COUNT(DISTINCT CASE
                WHEN mc.CurrentStatus = {1}
                AND mc.IsDeceased = 0
                AND DATEADD(year, 1, mc.JoinDate) <= GETDATE()
                THEN mc.PeopleId END) AS RetainedYear1,
            -- 3 year retention
            COUNT(DISTINCT CASE
                WHEN mc.CurrentStatus = {1}
                AND mc.IsDeceased = 0
                AND DATEADD(year, 3, mc.JoinDate) <= GETDATE()
                THEN mc.PeopleId END) AS RetainedYear3,
            -- 5 year retention
            COUNT(DISTINCT CASE
                WHEN mc.CurrentStatus = {1}
                AND mc.IsDeceased = 0
                AND DATEADD(year, 5, mc.JoinDate) <= GETDATE()
                THEN mc.PeopleId END) AS RetainedYear5
        FROM MemberCohorts mc
        GROUP BY mc.CohortYear
    )
    SELECT
        CohortYear,
        CohortSize,
        -- Rate columns are NULL until the full cohort has had time to mature.
        -- Without this gate, the most recent cohort would show a misleadingly
        -- low % because late joiners haven't yet crossed the 1Y / 3Y / 5Y line.
        CASE WHEN CohortSize > 0 AND DATEADD(year, 1, {4}) <= GETDATE()
             THEN RetainedYear1 ELSE NULL END AS RetainedYear1,
        CASE WHEN CohortSize > 0 AND DATEADD(year, 1, {4}) <= GETDATE()
             THEN (RetainedYear1 * 100.0 / CohortSize) ELSE NULL END AS RetentionRate1Year,
        CASE WHEN CohortSize > 0 AND DATEADD(year, 3, {4}) <= GETDATE()
             THEN RetainedYear3 ELSE NULL END AS RetainedYear3,
        CASE WHEN CohortSize > 0 AND DATEADD(year, 3, {4}) <= GETDATE()
             THEN (RetainedYear3 * 100.0 / CohortSize) ELSE NULL END AS RetentionRate3Year,
        CASE WHEN CohortSize > 0 AND DATEADD(year, 5, {4}) <= GETDATE()
             THEN RetainedYear5 ELSE NULL END AS RetainedYear5,
        CASE WHEN CohortSize > 0 AND DATEADD(year, 5, {4}) <= GETDATE()
             THEN (RetainedYear5 * 100.0 / CohortSize) ELSE NULL END AS RetentionRate5Year
    FROM RetentionData
    ORDER BY CohortYear DESC
    """.format(
        fiscal_year_sql,         # 0
        Config.MEMBER_STATUS_ID, # 1
        start_date,              # 2
        campus_filter(''),       # 3
        cohort_end               # 4
    )

    return q.QuerySql(sql)

def get_campus_breakdown(start_date, end_date):
    """Get membership breakdown by campus"""
    fiscal_year_sql = get_fiscal_year_sql().format("p.JoinDate")
    
    # Note: Campus table may not exist in all TouchPoint instances
    try:
        sql = """
        SELECT 
            {} AS JoinYear,
            ISNULL(c.Description, 'No Campus') AS Campus,
            COUNT(*) AS NewMembers,
            COUNT(DISTINCT p.FamilyId) AS NewFamilies
        FROM People p WITH (NOLOCK)
        LEFT JOIN Campus c ON p.CampusId = c.Id
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0
        GROUP BY {}, c.Description
        ORDER BY {} DESC, COUNT(*) DESC
        """.format(
            fiscal_year_sql,
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date,
            fiscal_year_sql,
            fiscal_year_sql
        )
        
        return q.QuerySql(sql)
    except:
        # Fallback if Campus table doesn't exist - just group by CampusId
        sql = """
        SELECT 
            {} AS JoinYear,
            CASE WHEN p.CampusId IS NULL THEN 'No Campus' 
                 ELSE 'Campus ' + CAST(p.CampusId AS varchar) END AS Campus,
            COUNT(*) AS NewMembers,
            COUNT(DISTINCT p.FamilyId) AS NewFamilies
        FROM People p WITH (NOLOCK)
        WHERE p.MemberStatusId = {}
          AND p.JoinDate >= '{}'
          AND p.JoinDate <= '{}'
          AND p.IsDeceased = 0
        GROUP BY {}, p.CampusId
        ORDER BY {} DESC, COUNT(*) DESC
        """.format(
            fiscal_year_sql,
            Config.MEMBER_STATUS_ID,
            start_date,
            end_date,
            fiscal_year_sql,
            fiscal_year_sql
        )
        
        return q.QuerySql(sql)

def get_lapsed_member_stats(start_date, end_date):
    """Get high-level lapsed/dropped member counts by fiscal year.
    Tracks people whose MemberStatusId changed FROM member to previous member."""
    fiscal_year_sql = get_fiscal_year_sql().format("cl.Created")
    sql = """
    SELECT
        {fy} AS LapsedYear,
        COUNT(DISTINCT cl.PeopleId) AS LapsedCount
    FROM ChangeLog cl WITH (NOLOCK)
    INNER JOIN People p WITH (NOLOCK) ON cl.PeopleId = p.PeopleId
    WHERE cl.Field = 'MemberStatusId'
      AND cl.Before = '{member}'
      AND cl.After = '{previous}'
      AND cl.Created >= '{start}'
      AND cl.Created <= '{end}'{campus}
    GROUP BY {fy}
    ORDER BY {fy} DESC
    """.format(
        fy=fiscal_year_sql,
        member=Config.MEMBER_STATUS_ID,
        previous=Config.PREVIOUS_MEMBER_STATUS_ID,
        start=start_date,
        end=end_date,
        campus=campus_filter()
    )
    try:
        return q.QuerySql(sql)
    except:
        return []

def get_status_transitions(start_date, end_date):
    """Get high-level member status transition summary by fiscal year.
    Shows flows between status categories: joined, dropped, returned."""
    fiscal_year_sql = get_fiscal_year_sql().format("cl.Created")
    sql = """
    SELECT
        {fy} AS TransYear,
        -- Became members (any status -> member)
        COUNT(DISTINCT CASE WHEN cl.After = '{member}'
            AND (cl.Before != '{member}' OR cl.Before IS NULL) THEN cl.PeopleId END) AS BecameMember,
        -- Left membership (member -> previous)
        COUNT(DISTINCT CASE WHEN cl.Before = '{member}'
            AND cl.After = '{previous}' THEN cl.PeopleId END) AS BecamePrevious,
        -- Returned to membership (previous -> member)
        COUNT(DISTINCT CASE WHEN cl.Before = '{previous}'
            AND cl.After = '{member}' THEN cl.PeopleId END) AS Returned,
        -- Net change
        COUNT(DISTINCT CASE WHEN cl.After = '{member}'
            AND (cl.Before != '{member}' OR cl.Before IS NULL) THEN cl.PeopleId END)
        - COUNT(DISTINCT CASE WHEN cl.Before = '{member}'
            AND cl.After = '{previous}' THEN cl.PeopleId END) AS NetChange
    FROM ChangeLog cl WITH (NOLOCK)
    INNER JOIN People p WITH (NOLOCK) ON cl.PeopleId = p.PeopleId
    WHERE cl.Field = 'MemberStatusId'
      AND cl.Created >= '{start}'
      AND cl.Created <= '{end}'{campus}
    GROUP BY {fy}
    ORDER BY {fy} DESC
    """.format(
        fy=fiscal_year_sql,
        member=Config.MEMBER_STATUS_ID,
        previous=Config.PREVIOUS_MEMBER_STATUS_ID,
        start=start_date,
        end=end_date,
        campus=campus_filter()
    )
    try:
        return q.QuerySql(sql)
    except:
        return []

def get_current_totals():
    """Get current membership totals"""
    sql = """
    SELECT 
        COUNT(*) AS TotalMembers,
        COUNT(DISTINCT FamilyId) AS TotalFamilies,
        COUNT(DISTINCT CampusId) AS TotalCampuses,
        AVG(Age) AS AverageAge,
        SUM(CASE WHEN GenderId = 1 THEN 1 ELSE 0 END) AS Males,
        SUM(CASE WHEN GenderId = 2 THEN 1 ELSE 0 END) AS Females
    FROM People WITH (NOLOCK)
    WHERE MemberStatusId = {}
      AND IsDeceased = 0
    """.format(Config.MEMBER_STATUS_ID)
    
    if Config.CAMPUS_ID > 0:
        sql += " AND CampusId = {}".format(Config.CAMPUS_ID)
    
    return q.QuerySqlTop1(sql)

def generate_report(analytics, start_date, end_date):
    """Generate the HTML report with all visualizations"""
    
    # Build campus dropdown options
    _campus_options = '<option value="0"{}>All Campuses</option>'.format(' selected' if Config.CAMPUS_ID == 0 else '')
    try:
        _campuses = q.QuerySql("SELECT Id, Description FROM lookup.Campus WHERE Id > 0 ORDER BY Description")
        for _c in _campuses:
            _cid = int(safe_get_value(_c, 'Id', 0))
            _cname = str(safe_get_value(_c, 'Description', 'Campus ' + str(_cid)))
            _sel = ' selected' if _cid == Config.CAMPUS_ID else ''
            _campus_options += '<option value="{}"{}>{}</option>'.format(_cid, _sel, _cname)
    except:
        pass

    # Report header with campus selector and settings gear
    print("""
    <div style="max-width:100%;overflow-x:hidden">
        <div class="row">
            <div class="col-md-12">
                <div class="page-header" style="position:relative">
                    <h1>{title} <small>{years} Year Main Analysis | {cohort} Year Cohort Analysis</small></h1>
                    <div style="display:flex;align-items:center;justify-content:space-between;margin-top:-5px">
                        <p class="text-muted" style="margin:0">Report Period: {start} to {end}</p>
                        <div style="display:flex;align-items:center;gap:12px">
                            <select id="marCampusSwitch" onchange="marSwitchCampus(this.value)"
                                    style="padding:4px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;background:#fff">
                                {campus_options}
                            </select>
                            <a href="/PyScriptForm/TPxi_MembershipAnalysisReport?settings=1"
                               style="font-size:14px;color:#999;text-decoration:none;white-space:nowrap"
                               title="Report Settings">
                                <i class="fa fa-cog"></i> Settings
                            </a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <script>
        function marSwitchCampus(campusId) {{
            var url = window.location.pathname;
            if (campusId && campusId !== '0') {{
                url += '?campus=' + campusId;
            }}
            window.location.href = url;
        }}
        </script>
    """.format(
        title=Config.REPORT_TITLE,
        campus_options=_campus_options,
        years=Config.YEARS_TO_ANALYZE,
        cohort=Config.COHORT_YEARS_TO_ANALYZE,
        start=format_date_display(start_date),
        end=format_date_display(end_date)
    ))
    
    try:
        # Executive Summary
        print('<script>console.log("Rendering executive summary...");</script>')
        render_executive_summary(analytics)
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering executive summary: {}</div>'.format(str(e)))
    
    try:
        # Yearly Trends Chart
        if 'yearly_trends' in analytics:
            print('<script>console.log("Rendering yearly trends...");</script>')
            render_yearly_trends_chart(analytics['yearly_trends'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering yearly trends: {}</div>'.format(str(e)))
    
    try:
        # Age Demographics Evolution
        if 'age_demographics' in analytics:
            print('<script>console.log("Rendering age demographics...");</script>')
            render_age_demographics_chart(analytics['age_demographics'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering age demographics: {}</div>'.format(str(e)))
    
    try:
        # Connection Pathways
        if Config.SHOW_ORIGIN_ANALYSIS and 'origins' in analytics:
            print('<script>console.log("Rendering origin analysis...");</script>')
            render_origin_analysis(analytics['origins'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering origin analysis: {}</div>'.format(str(e)))
    
    try:
        # Attendance Impact
        if Config.SHOW_ATTENDANCE_IMPACT and 'attendance_impact' in analytics:
            print('<script>console.log("Rendering attendance impact...");</script>')
            render_attendance_impact(analytics['attendance_impact'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering attendance impact: {}</div>'.format(str(e)))
    
    try:
        # Family Analysis
        if Config.SHOW_FAMILY_ANALYSIS and 'family_trends' in analytics:
            print('<script>console.log("Rendering family analysis...");</script>')
            render_family_analysis(analytics['family_trends'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering family analysis: {}</div>'.format(str(e)))
    
    try:
        # Retention Metrics
        if Config.SHOW_RETENTION_METRICS and 'retention' in analytics:
            print('<script>console.log("Rendering retention metrics...");</script>')
            render_retention_metrics(analytics['retention'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering retention metrics: {}</div>'.format(str(e)))
    
    try:
        # Baptism Age Bin Analysis
        if Config.SHOW_BAPTISM_ANALYSIS and 'baptism_age' in analytics:
            print('<script>console.log("Rendering baptism age analysis...");</script>')
            render_baptism_age_chart(analytics['baptism_age'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering baptism analysis: {}</div>'.format(str(e)))

    try:
        # Campus Breakdown
        if Config.SHOW_CAMPUS_BREAKDOWN and 'campus_breakdown' in analytics:
            print('<script>console.log("Rendering campus breakdown...");</script>')
            render_campus_breakdown(analytics['campus_breakdown'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering campus breakdown: {}</div>'.format(str(e)))

    try:
        # Status Transitions
        if 'transitions' in analytics and analytics['transitions']:
            print('<script>console.log("Rendering status transitions...");</script>')
            render_status_transitions(analytics['transitions'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering status transitions: {}</div>'.format(str(e)))

    try:
        # Lapsed Members
        if 'lapsed' in analytics and analytics['lapsed']:
            print('<script>console.log("Rendering lapsed members...");</script>')
            render_lapsed_members(analytics['lapsed'])
    except Exception as e:
        print('<div class="alert alert-warning">Error rendering lapsed members: {}</div>'.format(str(e)))

    print("</div>")  # Close container

def render_executive_summary(analytics):
    """Render executive summary section"""
    current = analytics.get('current_totals')
    yearly = analytics.get('yearly_trends', [])
    
    # Debug logging
    print('<script>console.log("Current totals type:", "{}");</script>'.format(type(current)))
    print('<script>console.log("Yearly trends count:", {});</script>'.format(len(yearly)))
    if current:
        try:
            # Try to see what attributes are available
            attrs = []
            for attr in dir(current):
                if not attr.startswith('_'):
                    attrs.append(attr)
            print('<script>console.log("Current attributes:", {});</script>'.format(str(attrs[:10])))
        except Exception as e:
            print('<script>console.log("Error getting attributes:", "{}");</script>'.format(str(e)))
    
    # Handle case where queries failed
    if not current:
        current = type('obj', (object,), {
            'TotalMembers': 0,
            'TotalFamilies': 0,
            'TotalCampuses': 0,
            'AverageAge': 0,
            'Males': 0,
            'Females': 0
        })()
    
    # Safe access to properties
    total_members = int(safe_get_value(current, 'TotalMembers', 0))
    
    # Calculate key metrics safely
    total_joined_years = 0
    try:
        if yearly:
            # Convert to list if it's not already
            yearly_list = list(yearly) if not isinstance(yearly, list) else yearly
            for y in yearly_list:
                total_joined_years += int(safe_get_value(y, 'NewMembers', 0))
    except Exception as e:
        print('<script>console.log("Error calculating total joined:", "{}");</script>'.format(str(e)))
        total_joined_years = 0
        
    avg_per_year = total_joined_years / len(yearly) if len(yearly) > 0 else 0
    
    # Calculate members joined this year and last year through same date
    current_date = model.DateTime
    current_fy_start_month = Config.FISCAL_YEAR_START_MONTH
    current_fy_start_day = Config.FISCAL_YEAR_START_DAY
    
    # Determine current fiscal year dates
    if Config.USE_FISCAL_YEAR:
        if current_date.Month >= current_fy_start_month:
            current_fy = current_date.Year + 1
            fy_start_date = "{}-{:02d}-{:02d}".format(current_date.Year, current_fy_start_month, current_fy_start_day)
        else:
            current_fy = current_date.Year
            fy_start_date = "{}-{:02d}-{:02d}".format(current_date.Year - 1, current_fy_start_month, current_fy_start_day)
    else:
        current_fy = current_date.Year
        fy_start_date = "{}-01-01".format(current_date.Year)
    
    # Get actual YTD numbers from database for accuracy
    joined_this_year_sql = """
        SELECT COUNT(*) FROM People 
        WHERE MemberStatusId = {} 
          AND JoinDate >= '{}'
          AND JoinDate <= GETDATE()
          AND IsDeceased = 0
    """.format(Config.MEMBER_STATUS_ID, fy_start_date)
    
    # Last year through same date
    if Config.USE_FISCAL_YEAR:
        last_fy_start = "{}-{:02d}-{:02d}".format(
            current_date.Year - 1 if current_date.Month >= current_fy_start_month else current_date.Year - 2,
            current_fy_start_month, 
            current_fy_start_day
        )
    else:
        last_fy_start = "{}-01-01".format(current_date.Year - 1)
    
    last_year_same_date_sql = """
        SELECT COUNT(*) FROM People 
        WHERE MemberStatusId = {} 
          AND JoinDate >= '{}'
          AND JoinDate <= DATEADD(year, -1, GETDATE())
          AND IsDeceased = 0
    """.format(Config.MEMBER_STATUS_ID, last_fy_start)
    
    joined_this_year = 0
    joined_last_year_ytd = 0
    
    try:
        joined_this_year = int(q.QuerySqlScalar(joined_this_year_sql) or 0)
        joined_last_year_ytd = int(q.QuerySqlScalar(last_year_same_date_sql) or 0)
        print('<script>console.log("Joined this FY: {}, Last FY same date: {}");</script>'.format(
            joined_this_year, joined_last_year_ytd))
    except Exception as e:
        print('<script>console.log("Error getting YTD numbers: {}");</script>'.format(str(e).replace('"', '')))
    
    # Calculate growth percentage
    if joined_last_year_ytd > 0:
        growth_pct = ((float(joined_this_year) / float(joined_last_year_ytd)) - 1) * 100
    else:
        growth_pct = 100 if joined_this_year > 0 else 0
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Executive Summary</h2>
            <div class="row">
                <div class="col-md-3">
                    <div class="panel panel-primary">
                        <div class="panel-body text-center">
                            <h2>{}</h2>
                            <p>Current Total Members</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="panel panel-success">
                        <div class="panel-body text-center">
                            <h2>{}</h2>
                            <p>Joined This {} YTD</p>
                            <small style="font-size: 10px; color: #666;">Through today's date</small>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="panel panel-info">
                        <div class="panel-body text-center">
                            <h2>{}</h2>
                            <p>Last {} Same Period</p>
                            <small style="font-size: 10px; color: #666;">Through same date last year</small>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="panel panel-{}">
                        <div class="panel-body text-center">
                            <h2>{}{}%</h2>
                            <p>YTD Growth</p>
                            <small style="font-size: 10px; color: #666;">Year-over-year comparison</small>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    """.format(
        format(total_members, ',d'),
        format(joined_this_year, ',d'),
        "FY" if Config.USE_FISCAL_YEAR else "Year",
        format(joined_last_year_ytd, ',d'),
        "FY" if Config.USE_FISCAL_YEAR else "Year",
        "success" if growth_pct >= 0 else "danger",
        "+" if growth_pct >= 0 else "",
        int(round(growth_pct))
    ))

def render_yearly_trends_chart(yearly_trends):
    """Render yearly membership trends chart"""
    # Reverse the data for charts (keep tables in DESC order but charts in ASC for chronological view)
    try:
        yearly_trends_chart = list(yearly_trends) if yearly_trends else []
        yearly_trends_chart.reverse()
    except:
        yearly_trends_chart = yearly_trends
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Membership Growth Trends</h2>
            <div class="panel panel-default">
                <div class="panel-body">
                    <canvas id="yearlyTrendsChart" style="max-height: 400px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
    var ctx = document.getElementById('yearlyTrendsChart').getContext('2d');
    var yearlyTrendsChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}{}'".format(get_year_label(), safe_get_value(y, 'JoinYear', '')) for y in yearly_trends_chart]) + """],
            datasets: [{
                label: 'New Members',
                data: [""" + ",".join([str(safe_get_value(y, 'NewMembers', 0)) for y in yearly_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[0] + """',
                backgroundColor: '""" + Config.CHART_COLORS[0] + """20',
                tension: 0.1,
                fill: true
            }, {
                label: 'New Families',
                data: [""" + ",".join([str(safe_get_value(y, 'NewFamilies', 0)) for y in yearly_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[1] + """',
                backgroundColor: '""" + Config.CHART_COLORS[1] + """20',
                tension: 0.1,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            elements: {
                line: {
                    borderWidth: 3
                },
                point: {
                    radius: 5,
                    hoverRadius: 7,
                    backgroundColor: 'white',
                    borderWidth: 2
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: 'New Members and Families by Year'
                },
                tooltip: {
                    mode: 'index',
                    intersect: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        }
    });
    </script>
    """)
    
    # Gender breakdown table
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Gender and Marital Status Breakdown</h3>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>New Members</th>
                            <th>New Families</th>
                            <th>Male</th>
                            <th>Female</th>
                            <th>% Male</th>
                            <th>Married</th>
                            <th>Single</th>
                            <th>Avg Age</th>
                        </tr>
                    </thead>
                    <tbody>
    """)
    
    # Build list for YoY comparison (data comes DESC, so index+1 is previous year)
    _trend_list = list(yearly_trends) if yearly_trends else []
    for i, year in enumerate(_trend_list):
        males = safe_get_value(year, 'Males', 0)
        new_members = safe_get_value(year, 'NewMembers', 0)
        male_pct = (float(males) / float(new_members) * 100) if new_members > 0 else 0
        avg_age = safe_get_value(year, 'AvgAgeAtJoining', None)

        # YoY delta indicator (compare to next item which is the prior year in DESC order)
        _yoy = ''
        if i < len(_trend_list) - 1:
            _prev = safe_get_value(_trend_list[i + 1], 'NewMembers', 0)
            if _prev > 0:
                _pct = int(round((float(new_members) - float(_prev)) / float(_prev) * 100))
                if _pct > 0:
                    _yoy = ' <span style="color:#27ae60;font-size:11px" title="vs prior year"><i class="fa fa-arrow-up"></i> +{}%</span>'.format(_pct)
                elif _pct < 0:
                    _yoy = ' <span style="color:#e74c3c;font-size:11px" title="vs prior year"><i class="fa fa-arrow-down"></i> {}%</span>'.format(_pct)
                else:
                    _yoy = ' <span style="color:#999;font-size:11px" title="vs prior year"><i class="fa fa-minus"></i> 0%</span>'

        print("""
                        <tr>
                            <td>{}</td>
                            <td>{}{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                        </tr>
        """.format(
            get_year_label() + str(safe_get_value(year, 'JoinYear', '')),
            new_members, _yoy,
            safe_get_value(year, 'NewFamilies', 0),
            males,
            safe_get_value(year, 'Females', 0),
            int(round(male_pct)),
            safe_get_value(year, 'Married', 0),
            safe_get_value(year, 'Single', 0),
            int(avg_age) if avg_age else "N/A"
        ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_age_demographics_chart(age_demographics):
    """Render age demographics evolution chart"""
    # Debug logging
    print('<script>console.log("Age demographics count:", {});</script>'.format(len(age_demographics) if age_demographics else 0))
    
    # Reverse for chronological chart display
    try:
        age_demographics_chart = list(age_demographics) if age_demographics else []
        age_demographics_chart.reverse()
    except:
        age_demographics_chart = age_demographics
    
    # Prepare data for stacked bar chart
    age_groups = []
    datasets = []
    
    # Get all age group names from Config
    # This is more reliable than trying to extract from the query results
    for group_name, min_age, max_age in Config.AGE_GROUPS:
        age_groups.append(group_name.replace(" ", ""))
    age_groups.append("UnknownAge")
    
    print('<script>console.log("Age groups:", {});</script>'.format(str(age_groups)))
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Age Demographics Evolution</h2>
            <div class="panel panel-default">
                <div class="panel-body">
                    <canvas id="ageDemographicsChart" style="max-height: 500px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    var ctx2 = document.getElementById('ageDemographicsChart').getContext('2d');
    var ageDemographicsChart = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: [""" + ",".join(["'{}{}'".format(get_year_label(), safe_get_value(y, 'JoinYear', '')) for y in age_demographics_chart]) + """],
            datasets: [
    """)
    
    # Create datasets for each age group
    for i, group in enumerate(age_groups):
        if i < len(Config.CHART_COLORS):
            color = Config.CHART_COLORS[i]
        else:
            color = "#999999"
        
        data_values = []
        for year in age_demographics_chart:
            value = safe_get_value(year, group, 0)
            data_values.append(str(value))
        
        # Format the group name for display
        display_name = group.replace("_", " ")
        if group == "UnknownAge":
            display_name = "Unknown Age"
        
        print("""
            {
                label: '""" + display_name + """',
                data: [""" + ",".join(data_values) + """],
                backgroundColor: '""" + color + """',
                borderColor: '""" + color + """',
                borderWidth: 1
            },
        """)
    
    print("""
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Age Distribution of New Members by Year',
                    font: {
                        size: 16
                    }
                },
                tooltip: {
                    mode: 'index',
                    intersect: false
                },
                legend: {
                    position: 'top'
                }
            },
            scales: {
                x: {
                    stacked: true,
                    title: {
                        display: true,
                        text: 'Fiscal Year'
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    },
                    title: {
                        display: true,
                        text: 'Number of New Members'
                    }
                }
            }
        }
    });
    </script>
    """)
    
    # Add a percentage breakdown table for clarity
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Age Distribution Percentages</h3>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>Total</th>
                            <th>Children (0-12)</th>
                            <th>Teens (13-17)</th>
                            <th>Young Adults (18-29)</th>
                            <th>Adults (30-49)</th>
                            <th>Mature Adults (50-64)</th>
                            <th>Seniors (65+)</th>
                        </tr>
                    </thead>
                    <tbody>
    """)
    
    for year in age_demographics:
        total = 0
        values = {}
        for group_name, min_age, max_age in Config.AGE_GROUPS:
            key = group_name.replace(" ", "")
            values[key] = safe_get_value(year, key, 0)
            total += values[key]
        
        if total > 0:
            print("""
                        <tr>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                            <td>{}%</td>
                            <td>{}%</td>
                            <td>{}%</td>
                            <td>{}%</td>
                            <td>{}%</td>
                        </tr>
            """.format(
                get_year_label() + str(safe_get_value(year, 'JoinYear', '')),
                total,
                int(round(values.get('Children', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('Teens', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('YoungAdults', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('Adults', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('MatureAdults', 0) * 100.0 / total)) if total > 0 else 0,
                int(round(values.get('Seniors', 0) * 100.0 / total)) if total > 0 else 0
            ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_origin_analysis(origins):
    """Render analysis of how people found the church"""
    # Aggregate origins by type
    origin_totals = {}
    origin_by_year = {}
    
    for row in origins:
        origin = safe_get_value(row, 'Origin', 'Unknown')
        year = safe_get_value(row, 'JoinYear', '')
        count = safe_get_value(row, 'Count', 0)
        
        if origin not in origin_totals:
            origin_totals[origin] = 0
            origin_by_year[origin] = {}
        
        origin_totals[origin] += count
        origin_by_year[origin][year] = count
    
    # Sort origins by total count
    sorted_origins = sorted(origin_totals.items(), key=lambda x: x[1], reverse=True)
    top_origins = sorted_origins[:10]  # Top 10 origins
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Connection Pathways - How People Found Us</h2>
        </div>
    </div>
    <div class="row">
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Top 10 Connection Pathways (10 Year Total)</h3>
                </div>
                <div class="panel-body">
                    <canvas id="originPieChart"></canvas>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Connection Pathway Trends</h3>
                </div>
                <div class="panel-body">
                    <canvas id="originTrendsChart" height="200"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    // Pie chart for top origins
    var ctx3 = document.getElementById('originPieChart').getContext('2d');
    var originPieChart = new Chart(ctx3, {
        type: 'doughnut',
        data: {
            labels: [""" + ",".join(["'{}'".format(o[0]) for o in top_origins]) + """],
            datasets: [{
                data: [""" + ",".join([str(o[1]) for o in top_origins]) + """],
                backgroundColor: [""" + ",".join(["'{}'".format(Config.CHART_COLORS[i % len(Config.CHART_COLORS)]) for i in range(len(top_origins))]) + """]
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: {
                    display: true,
                    text: 'How New Members Found Us'
                },
                legend: {
                    position: 'right'
                }
            }
        }
    });
    
    // Line chart for trends
    var ctx4 = document.getElementById('originTrendsChart').getContext('2d');
    var originTrendsChart = new Chart(ctx4, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}'".format(y) for y in sorted(set([r.JoinYear for r in origins]))]) + """],
            datasets: [
    """)
    
    # Create datasets for top 5 origins
    for i, (origin, total) in enumerate(top_origins[:5]):
        years = sorted(set([r.JoinYear for r in origins]))
        data_points = []
        for year in years:
            value = origin_by_year[origin].get(year, 0)
            data_points.append(str(value))
        
        print("""
            {
                label: '""" + origin + """',
                data: [""" + ",".join(data_points) + """],
                borderColor: '""" + Config.CHART_COLORS[i % len(Config.CHART_COLORS)] + """',
                backgroundColor: '""" + Config.CHART_COLORS[i % len(Config.CHART_COLORS)] + """20',
                tension: 0.1
            },
        """)
    
    print("""
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Top 5 Connection Pathways Over Time'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        }
    });
    </script>
    """)

def render_attendance_impact(attendance_impact):
    """Render attendance impact analysis"""
    # Reverse for chronological display
    try:
        attendance_impact_chart = list(attendance_impact) if attendance_impact else []
        attendance_impact_chart.reverse()
    except:
        attendance_impact_chart = attendance_impact
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Member Engagement Analysis</h2>
            <p class="text-muted">Individual attendance tracking 6 months before and after membership. Shows engagement patterns and integration success.</p>
            <div class="alert alert-info">
                <strong>6-Month Analysis Periods:</strong>
                <ul class="mb-0">
                    <li><strong>Before:</strong> 6 months prior to joining date - shows pre-membership connection level</li>
                    <li><strong>After:</strong> 6 months following joining date - shows post-membership integration success</li>
                </ul>
            </div>
        </div>
    </div>
    <div class="row">
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">{} Engagement</h3>
                </div>
                <div class="panel-body">
                    <canvas id="worshipAttendanceChart" style="max-height: 300px;"></canvas>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Overall Ministry Engagement</h3>
                </div>
                <div class="panel-body">
                    <canvas id="smallGroupAttendanceChart" style="max-height: 300px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    """.format(Config.PRIMARY_WORSHIP_PROGRAM_NAME, Config.SMALL_GROUP_PROGRAM_NAME))
    
    # Create comparison charts
    print("""
    <script>
    // Worship headcount trend
    var ctx5 = document.getElementById('worshipAttendanceChart').getContext('2d');
    var worshipChart = new Chart(ctx5, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}{}'".format(get_year_label(), safe_get_value(y, 'JoinYear', '')) for y in attendance_impact_chart]) + """],
            datasets: [{
                label: 'Average Worship Headcount',
                data: [""" + ",".join([str(int(safe_get_value(y, 'AvgWorshipHeadcount', 0))) for y in attendance_impact_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[0] + """',
                backgroundColor: '""" + Config.CHART_COLORS[0] + """20',
                borderWidth: 3,
                tension: 0.1,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            elements: {
                point: {
                    radius: 5,
                    hoverRadius: 7
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: 'Worship Service Average Headcount by Year'
                }
            },
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
    
    // Small group attendance comparison
    var ctx6 = document.getElementById('smallGroupAttendanceChart').getContext('2d');
    var smallGroupChart = new Chart(ctx6, {
        type: 'bar',
        data: {
            labels: [""" + ",".join(["'{}{}'".format(get_year_label(), safe_get_value(y, 'JoinYear', '')) for y in attendance_impact_chart]) + """],
            datasets: [{
                label: 'Before Membership',
                data: [""" + ",".join([str(int(safe_get_value(y, 'AvgSmallGroupBefore', 0))) for y in attendance_impact_chart]) + """],
                backgroundColor: '""" + Config.CHART_COLORS[3] + """'
            }, {
                label: 'After Membership',
                data: [""" + ",".join([str(int(safe_get_value(y, 'AvgSmallGroupAfter', 0))) for y in attendance_impact_chart]) + """],
                backgroundColor: '""" + Config.CHART_COLORS[1] + """'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Average Connect Group Attendance (6 month periods)'
                }
            },
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
    </script>
    """)
    
    # Summary statistics
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Member Integration Success Metrics</h3>
            <div class="alert alert-info">
                <p><strong>Understanding the Metrics:</strong></p>
                <ul class="mb-0">
                    <li><strong>CG (Connect Groups):</strong> Specifically tracks attendance at Connect Group/Small Group meetings (Program ID: {})</li>
                    <li><strong>Ministry:</strong> Tracks attendance at ALL church programs and activities EXCEPT worship services - includes Connect Groups, classes, serve teams, events, committees, etc.</li>
                    <li><strong>Before/After:</strong> Average individual attendance counts in the 6-month periods before and after membership</li>
                    <li><strong>Improved:</strong> Number of people who had higher attendance after joining than before</li>
                    <li><strong>% Improved:</strong> Percentage of new members who increased their engagement level</li>
                </ul>
            </div>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>New Members</th>
                            <th>Avg CG Before</th>
                            <th>Avg CG After</th>
                            <th>CG Improved</th>
                            <th>% CG Improved</th>
                            <th>Avg Ministry Before</th>
                            <th>Avg Ministry After</th>
                            <th>Ministry Improved</th>
                            <th>% Ministry Improved</th>
                        </tr>
                    </thead>
                    <tbody>
    """.format(Config.SMALL_GROUP_PROGRAM_ID))
    
    for year in attendance_impact:
        total_members = safe_get_value(year, 'TotalMembers', 0)
        cg_before = safe_get_value(year, 'AvgSmallGroupBefore', 0)
        cg_after = safe_get_value(year, 'AvgSmallGroupAfter', 0)
        cg_improved = safe_get_value(year, 'ImprovedSmallGroupCount', 0)
        
        # Get ministry data from SQL results (now available in query)
        ministry_before = safe_get_value(year, 'AvgMinistryBefore', 0)
        ministry_after = safe_get_value(year, 'AvgMinistryAfter', 0)
        ministry_improved = safe_get_value(year, 'ImprovedMinistryCount', 0)
        
        cg_improved_pct = (float(cg_improved) / float(total_members) * 100) if total_members > 0 else 0
        ministry_improved_pct = (float(ministry_improved) / float(total_members) * 100) if total_members > 0 else 0
        
        print("""
                        <tr>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                        </tr>
        """.format(
            get_year_label() + str(safe_get_value(year, 'JoinYear', '')),
            total_members,
            round(cg_before * 10) / 10.0,
            round(cg_after * 10) / 10.0,
            cg_improved,
            int(round(cg_improved_pct)),
            round(ministry_before * 10) / 10.0,
            round(ministry_after * 10) / 10.0,
            ministry_improved,
            int(round(ministry_improved_pct))
        ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_family_analysis(family_trends):
    """Render family composition analysis"""
    # Reverse for chronological display
    try:
        family_trends_chart = list(family_trends) if family_trends else []
        family_trends_chart.reverse()
    except:
        family_trends_chart = family_trends
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Family Unit Analysis</h2>
        </div>
    </div>
    <div class="row">
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Family Size Trends</h3>
                </div>
                <div class="panel-body">
                    <canvas id="familySizeChart" style="width: 100%; height: 300px;"></canvas>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Families with/without Children</h3>
                </div>
                <div class="panel-body">
                    <canvas id="familyCompositionChart" style="width: 100%; height: 300px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    // Family size trends
    var ctx7 = document.getElementById('familySizeChart').getContext('2d');
    var familySizeChart = new Chart(ctx7, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}'".format(safe_get_value(y, 'JoinYear', '')) for y in family_trends_chart]) + """],
            datasets: [{
                label: 'Single Person',
                data: [""" + ",".join([str(safe_get_value(y, 'SinglePersonFamilies', 0)) for y in family_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[0] + """',
                tension: 0.1
            }, {
                label: 'Couples',
                data: [""" + ",".join([str(safe_get_value(y, 'CouplesFamilies', 0)) for y in family_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[1] + """',
                tension: 0.1
            }, {
                label: 'Small Families (3-4)',
                data: [""" + ",".join([str(safe_get_value(y, 'SmallFamilies', 0)) for y in family_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[2] + """',
                tension: 0.1
            }, {
                label: 'Large Families (5+)',
                data: [""" + ",".join([str(safe_get_value(y, 'LargeFamilies', 0)) for y in family_trends_chart]) + """],
                borderColor: '""" + Config.CHART_COLORS[3] + """',
                tension: 0.1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Family Size Trends'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        }
    });
    
    // Family composition
    var ctx8 = document.getElementById('familyCompositionChart').getContext('2d');
    var familyCompChart = new Chart(ctx8, {
        type: 'bar',
        data: {
            labels: [""" + ",".join(["'{}'".format(safe_get_value(y, 'JoinYear', '')) for y in family_trends_chart]) + """],
            datasets: [{
                label: 'Families with Children',
                data: [""" + ",".join([str(safe_get_value(y, 'FamiliesWithChildren', 0)) for y in family_trends_chart]) + """],
                backgroundColor: '""" + Config.CHART_COLORS[1] + """',
                borderColor: '""" + Config.CHART_COLORS[1] + """',
                borderWidth: 1
            }, {
                label: 'Families without Children',
                data: [""" + ",".join([str(safe_get_value(y, 'UniqueFamilies', 0) - safe_get_value(y, 'FamiliesWithChildren', 0)) for y in family_trends_chart]) + """],
                backgroundColor: '""" + Config.CHART_COLORS[4] + """',
                borderColor: '""" + Config.CHART_COLORS[4] + """',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Families with/without Children'
                }
            },
            scales: {
                x: {
                    stacked: true
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        }
    });
    </script>
    """)
    
    # Average family size trend
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Average Family Size Trend</h3>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>Unique Families</th>
                            <th>Average Family Size</th>
                            <th>% with Children</th>
                        </tr>
                    </thead>
                    <tbody>
    """)
    
    for year in family_trends:
        unique_families = safe_get_value(year, 'UniqueFamilies', 0)
        families_with_children = safe_get_value(year, 'FamiliesWithChildren', 0)
        children_pct = (float(families_with_children) / float(unique_families) * 100) if unique_families > 0 else 0
        
        print("""
                        <tr>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}%</td>
                        </tr>
        """.format(
            safe_get_value(year, 'JoinYear', ''),
            unique_families,
            int(round(float(safe_get_value(year, 'AvgFamilySize', 0)) * 10)) / 10.0,
            int(round(children_pct))
        ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_retention_metrics(retention_data):
    """Render retention metrics visualization"""
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Member Retention Analysis</h2>
            <p class="text-muted">Percentage of members from each year's cohort who remain active members</p>
        </div>
    </div>
    <div class="row">
        <div class="col-md-12">
            <div class="panel panel-default">
                <div class="panel-body">
                    <canvas id="retentionChart" style="max-height: 300px;"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    var ctx9 = document.getElementById('retentionChart').getContext('2d');
    var retentionChart = new Chart(ctx9, {
        type: 'line',
        data: {
            labels: [""" + ",".join(["'{}'".format(safe_get_value(y, 'CohortYear', '')) for y in retention_data if safe_get_value(y, 'RetentionRate1Year', None) is not None]) + """],
            datasets: [{
                label: '1 Year Retention',
                data: [""" + ",".join([str(int(safe_get_value(y, 'RetentionRate1Year', 0))) for y in retention_data if safe_get_value(y, 'RetentionRate1Year', None) is not None]) + """],
                borderColor: '""" + Config.CHART_COLORS[0] + """',
                backgroundColor: '""" + Config.CHART_COLORS[0] + """20',
                tension: 0.1
            }, {
                label: '3 Year Retention',
                data: [""" + ",".join([str(int(safe_get_value(y, 'RetentionRate3Year', 0))) if safe_get_value(y, 'RetentionRate3Year', None) is not None else "null" for y in retention_data if safe_get_value(y, 'RetentionRate1Year', None) is not None]) + """],
                borderColor: '""" + Config.CHART_COLORS[1] + """',
                backgroundColor: '""" + Config.CHART_COLORS[1] + """20',
                tension: 0.1,
                spanGaps: true
            }, {
                label: '5 Year Retention',
                data: [""" + ",".join([str(int(safe_get_value(y, 'RetentionRate5Year', 0))) if safe_get_value(y, 'RetentionRate5Year', None) is not None else "null" for y in retention_data if safe_get_value(y, 'RetentionRate1Year', None) is not None]) + """],
                borderColor: '""" + Config.CHART_COLORS[2] + """',
                backgroundColor: '""" + Config.CHART_COLORS[2] + """20',
                tension: 0.1,
                spanGaps: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            elements: {
                line: {
                    borderWidth: 3
                },
                point: {
                    radius: 5,
                    hoverRadius: 7,
                    backgroundColor: 'white',
                    borderWidth: 2
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: 'Member Retention Rates by Cohort Year'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + (context.parsed.y || 0) + '%';
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100,
                    ticks: {
                        callback: function(value) {
                            return value + '%';
                        }
                    }
                }
            }
        }
    });
    </script>
    """)
    
    # Detailed retention table
    # Year-type label switches with USE_FISCAL_YEAR setting -- legend used to
    # say "fiscal year" even when the user had Calendar Year selected.
    _year_type_label = "fiscal year" if Config.USE_FISCAL_YEAR else "calendar year"
    _example_year = "FY2020" if Config.USE_FISCAL_YEAR else "2020"
    _example_next = "FY2021" if Config.USE_FISCAL_YEAR else "2021"
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Retention Details by Cohort</h3>
            <div class="alert alert-info">
                <p><strong>Understanding Retention Metrics:</strong></p>
                <ul>
                    <li><strong>Cohort Year</strong> - The {0} when members joined</li>
                    <li><strong>Initial Size</strong> - Number of new members who joined that year</li>
                    <li><strong>1/3/5 Year columns</strong> - How many from that cohort are still active members after that many years</li>
                    <li><strong>Percentage columns</strong> - What percentage of the original cohort remains active</li>
                    <li>Dashes (-) indicate not enough time has passed for the whole cohort to mature for that metric</li>
                </ul>
                <p><em>Example: If 100 people joined in {1} and 85 are still members in {2}, the 1-year retention is 85%</em></p>
                <hr style="border-color:#a6cdef;margin:8px 0;">
                <p style="margin-bottom:4px;"><strong>What these rates actually measure (and what they don't):</strong></p>
                <p style="margin-bottom:0;font-size:0.95em;">
                    These columns measure <strong>snapshot retention</strong>, whether each cohort member
                    is <em>still a Member today</em>, gated by how many years have passed since they joined.
                    They do <strong>not</strong> measure point-in-time retention (i.e. whether someone was
                    still a Member exactly at the 1-year mark vs. the 3-year mark). That distinction matters
                    for older cohorts: if everyone who was going to leave has already left, the 1Y / 3Y / 5Y
                    columns will collapse to the same number. A 2020 cohort showing
                    <code>87% / 87% / 87%</code> means "of the people who joined in 2020, 87% are still
                    Members today".  The math can't distinguish whether the 13% who left did so in
                    year 1 or year 4. For true per-anniversary retention we'd need to mine
                    <code>ChangeLog</code> for historical status changes (not yet implemented).
                </p>
            </div>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Cohort Year</th>
                            <th>Initial Size</th>
                            <th>1 Year</th>
                            <th>1 Year %</th>
                            <th>3 Years</th>
                            <th>3 Years %</th>
                            <th>5 Years</th>
                            <th>5 Years %</th>
                        </tr>
                    </thead>
                    <tbody>
    """.format(_year_type_label, _example_year, _example_next))

    for cohort in retention_data:
        rate1 = safe_get_value(cohort, 'RetentionRate1Year', None)
        rate3 = safe_get_value(cohort, 'RetentionRate3Year', None)
        rate5 = safe_get_value(cohort, 'RetentionRate5Year', None)
        ret1 = safe_get_value(cohort, 'RetainedYear1', None)
        ret3 = safe_get_value(cohort, 'RetainedYear3', None)
        ret5 = safe_get_value(cohort, 'RetainedYear5', None)

        print("""
                        <tr>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                            <td>{}</td>
                        </tr>
        """.format(
            safe_get_value(cohort, 'CohortYear', ''),
            safe_get_value(cohort, 'CohortSize', 0),
            ret1 if ret1 is not None else "-",
            "{}%".format(int(round(rate1))) if rate1 is not None else "-",
            ret3 if ret3 is not None else "-",
            "{}%".format(int(round(rate3))) if rate3 is not None else "-",
            ret5 if ret5 is not None else "-",
            "{}%".format(int(round(rate5))) if rate5 is not None else "-"
        ))
    
    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_status_transitions(transitions):
    """Render high-level member status transition table with YoY indicators"""
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Membership Status Transitions</h2>
            <p class="text-muted">Tracks when people move between member status categories based on change log records.</p>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>Became Member</th>
                            <th>Left Membership</th>
                            <th>Returned</th>
                            <th>Net Change</th>
                        </tr>
                    </thead>
                    <tbody>
    """)
    _tlist = list(transitions) if transitions else []
    for i, row in enumerate(_tlist):
        _became = int(safe_get_value(row, 'BecameMember', 0))
        _left = int(safe_get_value(row, 'BecamePrevious', 0))
        _returned = int(safe_get_value(row, 'Returned', 0))
        _net = int(safe_get_value(row, 'NetChange', 0))

        # Net change indicator
        if _net > 0:
            _net_html = '<span style="color:#27ae60;font-weight:bold"><i class="fa fa-arrow-up"></i> +{}</span>'.format(_net)
        elif _net < 0:
            _net_html = '<span style="color:#e74c3c;font-weight:bold"><i class="fa fa-arrow-down"></i> {}</span>'.format(_net)
        else:
            _net_html = '<span style="color:#999">0</span>'

        print("""
                        <tr>
                            <td>{yr}</td>
                            <td><span style="color:#27ae60">{became}</span></td>
                            <td><span style="color:#e74c3c">{left}</span></td>
                            <td><span style="color:#3498db">{returned}</span></td>
                            <td>{net}</td>
                        </tr>
        """.format(
            yr=get_year_label() + str(safe_get_value(row, 'TransYear', '')),
            became=_became,
            left=_left,
            returned=_returned,
            net=_net_html
        ))

    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_lapsed_members(lapsed_data):
    """Render high-level lapsed member counts by fiscal year"""
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Lapsed Members</h2>
            <p class="text-muted">Members whose status changed from Member to Previous Member, by year.</p>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>Members Lost</th>
                            <th>Trend</th>
                        </tr>
                    </thead>
                    <tbody>
    """)
    _llist = list(lapsed_data) if lapsed_data else []
    for i, row in enumerate(_llist):
        _count = int(safe_get_value(row, 'LapsedCount', 0))

        # YoY trend
        _trend = ''
        if i < len(_llist) - 1:
            _prev = int(safe_get_value(_llist[i + 1], 'LapsedCount', 0))
            if _prev > 0:
                _pct = int(round((float(_count) - float(_prev)) / float(_prev) * 100))
                if _pct > 0:
                    # More lapsed = bad (red up arrow)
                    _trend = '<span style="color:#e74c3c;font-size:11px"><i class="fa fa-arrow-up"></i> +{}% vs prior</span>'.format(_pct)
                elif _pct < 0:
                    # Fewer lapsed = good (green down arrow)
                    _trend = '<span style="color:#27ae60;font-size:11px"><i class="fa fa-arrow-down"></i> {}% vs prior</span>'.format(_pct)
                else:
                    _trend = '<span style="color:#999;font-size:11px">No change</span>'

        print("""
                        <tr>
                            <td>{yr}</td>
                            <td>{count}</td>
                            <td>{trend}</td>
                        </tr>
        """.format(
            yr=get_year_label() + str(safe_get_value(row, 'LapsedYear', '')),
            count=_count,
            trend=_trend
        ))

    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

def render_campus_breakdown(campus_data):
    """Render campus-specific membership analysis"""
    # Aggregate by campus
    campus_totals = {}
    for row in campus_data:
        campus = safe_get_value(row, 'Campus', 'Unknown')
        if campus not in campus_totals:
            campus_totals[campus] = 0
        campus_totals[campus] += safe_get_value(row, 'NewMembers', 0)
    
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Multi-Campus Analysis</h2>
        </div>
    </div>
    <div class="row">
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">10-Year Total by Campus</h3>
                </div>
                <div class="panel-body">
                    <canvas id="campusPieChart"></canvas>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Campus Growth Trends</h3>
                </div>
                <div class="panel-body">
                    <canvas id="campusTrendsChart" height="200"></canvas>
                </div>
            </div>
        </div>
    </div>
    
    <script>
    // Campus pie chart
    var ctx10 = document.getElementById('campusPieChart').getContext('2d');
    var campusPieChart = new Chart(ctx10, {
        type: 'pie',
        data: {
            labels: [""" + ",".join(["'{}'".format(c) for c in campus_totals.keys()]) + """],
            datasets: [{
                data: [""" + ",".join([str(v) for v in campus_totals.values()]) + """],
                backgroundColor: [""" + ",".join(["'{}'".format(Config.CHART_COLORS[i % len(Config.CHART_COLORS)]) for i in range(len(campus_totals))]) + """]
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: {
                    display: true,
                    text: 'Total New Members by Campus (10 Years)'
                }
            }
        }
    });
    </script>
    """)

def get_baptism_age_trends(start_date, end_date):
    """Get baptism counts by age bin and fiscal year.
    Calculates age at time of baptism using BDate and BaptismDate.
    Falls back to current Age if BDate is not available.
    """
    fiscal_year_sql = get_fiscal_year_sql().format("p.BaptismDate")

    # Calculate age at baptism: DATEDIFF(year, BDate, BaptismDate) adjusted for birthday
    # Falls back to current p.Age if BDate is null
    age_at_baptism = """CASE
            WHEN p.BDate IS NOT NULL THEN
                DATEDIFF(year, p.BDate, p.BaptismDate)
                - CASE WHEN DATEADD(year, DATEDIFF(year, p.BDate, p.BaptismDate), p.BDate) > p.BaptismDate THEN 1 ELSE 0 END
            ELSE p.Age
        END"""

    # Build age bin CASE expression using age at baptism
    age_case = "CASE\n"
    for label, min_age, max_age in Config.BAPTISM_AGE_BINS:
        if max_age >= 150:
            age_case += "            WHEN ({}) >= {} THEN '{}'\n".format(age_at_baptism, min_age, label)
        else:
            age_case += "            WHEN ({}) >= {} AND ({}) <= {} THEN '{}'\n".format(age_at_baptism, min_age, age_at_baptism, max_age, label)
    age_case += "            WHEN ({}) IS NULL THEN 'Unknown'\n".format(age_at_baptism)
    age_case += "        END"

    sql = """
    SELECT
        {fiscal_year} AS BaptismYear,
        {age_case} AS AgeBin,
        COUNT(*) AS BaptismCount
    FROM People p WITH (NOLOCK)
    WHERE p.BaptismDate IS NOT NULL
      AND p.BaptismDate >= '{start}'
      AND p.BaptismDate <= '{end}'
      AND p.IsDeceased = 0{campus}
    GROUP BY {fiscal_year}, {age_case}
    ORDER BY {fiscal_year} DESC, {age_case}
    """.format(
        fiscal_year=fiscal_year_sql,
        age_case=age_case,
        start=start_date,
        end=end_date,
        campus=campus_filter()
    )

    return q.QuerySql(sql)


def render_baptism_age_chart(baptism_data):
    """Render baptism age bin breakdown by year as stacked bar chart and table"""
    if not baptism_data or len(baptism_data) == 0:
        print("""
        <div class="row">
            <div class="col-md-12">
                <h2>Baptism Age Analysis</h2>
                <div class="alert alert-info">No baptism data found for the selected period.</div>
            </div>
        </div>
        """)
        return

    # Pivot data: collect years and age bins
    years_set = {}
    bins_set = []
    for row in baptism_data:
        yr = safe_get_value(row, 'BaptismYear', '')
        ab = safe_get_value(row, 'AgeBin', 'Unknown')
        cnt = safe_get_value(row, 'BaptismCount', 0)
        if yr not in years_set:
            years_set[yr] = {}
        years_set[yr][ab] = cnt
        if ab not in bins_set:
            bins_set.append(ab)

    # Sort years ascending for chart
    sorted_years = sorted(years_set.keys())

    # Use config bin order, then append any extras (like Unknown)
    ordered_bins = [label for label, _, _ in Config.BAPTISM_AGE_BINS]
    if 'Unknown' in bins_set:
        ordered_bins.append('Unknown')

    # Build chart labels
    year_labels = ",".join(["'{}{}'".format(get_year_label(), y) for y in sorted_years])

    # Chart colors
    bin_colors = [
        "#667eea", "#48bb78", "#ed8936", "#e53e3e", "#38b2ac",
        "#805ad5", "#d69e2e", "#3182ce", "#e84393", "#999999"
    ]

    # Short labels for table columns, full labels for chart legend
    short_labels = {}
    for label, min_age, max_age in Config.BAPTISM_AGE_BINS:
        if 'Preschool' in label:
            short_labels[label] = 'PS'
        elif 'Children' in label:
            short_labels[label] = 'CH'
        elif 'Preteen' in label:
            short_labels[label] = 'PT'
        elif 'Middle' in label:
            short_labels[label] = 'MS'
        elif 'High' in label:
            short_labels[label] = 'HS'
        else:
            short_labels[label] = label
    short_labels['Unknown'] = 'Unk'

    # Start chart section
    print("""
    <div class="row">
        <div class="col-md-12">
            <h2>Baptism Age Analysis</h2>
            <div class="panel panel-default">
                <div class="panel-body">
                    <div style="height: 400px;">
                        <canvas id="baptismAgeChart"></canvas>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
    var ctxBaptism = document.getElementById('baptismAgeChart').getContext('2d');
    var baptismAgeChart = new Chart(ctxBaptism, {
        type: 'bar',
        data: {
            labels: [""" + year_labels + """],
            datasets: [
    """)

    # Create a dataset for each age bin
    for i, bin_label in enumerate(ordered_bins):
        color = bin_colors[i % len(bin_colors)]
        data_values = []
        for yr in sorted_years:
            data_values.append(str(years_set.get(yr, {}).get(bin_label, 0)))

        print("""
            {{
                label: '{}',
                data: [{}],
                backgroundColor: '{}',
                borderColor: '{}',
                borderWidth: 1
            }},
        """.format(bin_label, ",".join(data_values), color, color))

    print("""
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Baptisms by Age Group per Year',
                    font: { size: 16 }
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    callbacks: {
                        footer: function(tooltipItems) {
                            var total = 0;
                            tooltipItems.forEach(function(item) { total += item.parsed.y; });
                            return 'Total: ' + total;
                        }
                    }
                },
                legend: {
                    position: 'top',
                    labels: {
                        boxWidth: 12,
                        padding: 10,
                        font: { size: 11 }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    title: {
                        display: true,
                        text: '""" + ("Fiscal Year" if Config.USE_FISCAL_YEAR else "Year") + """'
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: { precision: 0 },
                    title: {
                        display: true,
                        text: 'Number of Baptisms'
                    }
                }
            }
        }
    });
    </script>
    """)

    # Render data table with short column headers and clickable numbers
    print("""
    <div class="row">
        <div class="col-md-12">
            <h3>Baptism Age Breakdown by Year</h3>
            <p class="text-muted"><small>Click any number to see the people in that group.</small></p>
            <div class="table-responsive">
                <table class="table table-striped table-condensed">
                    <thead>
                        <tr>
                            <th>Year</th>
                            <th>Total</th>
    """)
    for bin_label in ordered_bins:
        print('<th style="text-align:center">{}</th>'.format(short_labels.get(bin_label, bin_label)))
    print("""
                        </tr>
                    </thead>
                    <tbody>
    """)

    # Table rows in descending year order
    for yr in sorted(years_set.keys(), reverse=True):
        yr_data = years_set[yr]
        total = sum(yr_data.get(b, 0) for b in ordered_bins)
        print('<tr>')
        print('<td>{}{}</td>'.format(get_year_label(), yr))
        print('<td><strong>{}</strong></td>'.format(total))
        for bin_label in ordered_bins:
            count = yr_data.get(bin_label, 0)
            pct = int(round(count * 100.0 / total)) if total > 0 else 0
            if count > 0:
                print('<td style="text-align:center"><a href="javascript:void(0)" onclick="baptismDrill({}, \'{}\')" style="cursor:pointer;color:#667eea;text-decoration:underline">{}</a> ({}%)</td>'.format(yr, bin_label.replace("'", "\\'"), count, pct))
            else:
                print('<td style="text-align:center">0 (0%)</td>')
        print('</tr>')

    print("""
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """)

    # Age group key/legend
    print("""
    <div class="row">
        <div class="col-md-12">
            <div class="panel panel-default">
                <div class="panel-heading"><strong>Age Group Key</strong></div>
                <div class="panel-body">
                    <table class="table table-condensed" style="margin-bottom:0">
                        <tbody>
                            <tr><td style="width:60px"><strong>PS</strong></td><td>Preschool (0-4)</td></tr>
                            <tr><td><strong>CH</strong></td><td>Children (5-9)</td></tr>
                            <tr><td><strong>PT</strong></td><td>Preteen (10-11)</td></tr>
                            <tr><td><strong>MS</strong></td><td>Middle School (12-14)</td></tr>
                            <tr><td><strong>HS</strong></td><td>High School (15-18)</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    """)

    # Modal HTML and JavaScript for drill-down
    print("""
    <!-- Baptism Drill-Down Modal -->
    <div id="baptismModal" style="display:none; position:fixed; z-index:9999; left:0; top:0; width:100%; height:100%; background-color:rgba(0,0,0,0.5); overflow:auto;">
        <div style="background-color:#fff; margin:50px auto; padding:0; border:1px solid #888; width:90%; max-width:700px; border-radius:8px; box-shadow:0 20px 60px rgba(0,0,0,0.3);">
            <div id="baptismModalHeader" style="padding:12px 20px; border-bottom:1px solid #e5e7eb; display:flex; justify-content:space-between; align-items:center; background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius:8px 8px 0 0;">
                <span id="baptismModalTitle" style="color:white; font-size:16px; font-weight:bold;"></span>
                <span onclick="closeBaptismModal()" style="color:white; font-size:24px; font-weight:bold; cursor:pointer; line-height:1;">&times;</span>
            </div>
            <div id="baptismModalContent" style="padding:15px; max-height:500px; overflow-y:auto;">
                <div style="text-align:center; padding:20px;"><i class="fa fa-spinner fa-spin"></i> Loading...</div>
            </div>
        </div>
    </div>

    <script>
    function baptismDrill(year, bin) {
        var modal = document.getElementById('baptismModal');
        var content = document.getElementById('baptismModalContent');
        var title = document.getElementById('baptismModalTitle');

        title.textContent = 'FY' + year + ' - ' + bin;
        content.innerHTML = '<div style="text-align:center; padding:20px;"><i class="fa fa-spinner fa-spin"></i> Loading...</div>';
        modal.style.display = 'block';

        var baseUrl = window.location.pathname;
        var url = baseUrl + '?baptism_drill=true&yr=' + encodeURIComponent(year) + '&bin=' + encodeURIComponent(bin);

        var xhr = new XMLHttpRequest();
        xhr.open('GET', url, true);
        xhr.onload = function() {
            if (xhr.status === 200) {
                var resp = xhr.responseText;
                var startMarker = '<!-- AJAX_CONTENT_START -->';
                var endMarker = '<!-- AJAX_CONTENT_END -->';
                var s = resp.indexOf(startMarker);
                var e = resp.indexOf(endMarker);
                if (s !== -1 && e !== -1) {
                    content.innerHTML = resp.substring(s + startMarker.length, e);
                } else {
                    content.innerHTML = '<div class="alert alert-danger">Could not parse response.</div>';
                }
            } else {
                content.innerHTML = '<div class="alert alert-danger">Error loading data.</div>';
            }
        };
        xhr.onerror = function() {
            content.innerHTML = '<div class="alert alert-danger">Network error.</div>';
        };
        xhr.send();
    }

    function closeBaptismModal() {
        document.getElementById('baptismModal').style.display = 'none';
    }

    // Close modal on Escape key or clicking outside
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') closeBaptismModal();
    });
    document.getElementById('baptismModal').addEventListener('click', function(e) {
        if (e.target === this) closeBaptismModal();
    });
    </script>
    """)



def format_date_display(date_str):
    """Format date for display"""
    try:
        # Parse the date string and format nicely
        date_obj = model.ParseDate(date_str[:10])
        return date_obj.ToString("MMMM d, yyyy")
    except:
        return date_str[:10]

def safe_get_value(obj, attr, default=0):
    """Safely get value from TouchPoint row object"""
    try:
        if hasattr(obj, attr):
            val = getattr(obj, attr)
            if val is not None:
                return val
    except:
        pass
    return default

def get_fiscal_year_sql():
    """Get SQL expression to calculate fiscal year from a date"""
    if not Config.USE_FISCAL_YEAR:
        return "YEAR({0})"  # Calendar year
    
    # For fiscal year starting Oct 1: dates from Oct-Dec are in next fiscal year
    # Example: Oct 1, 2023 is in fiscal year 2024
    if Config.FISCAL_YEAR_START_MONTH >= 10:  # Oct, Nov, Dec
        return """
        CASE 
            WHEN MONTH({0}) >= {1} THEN YEAR({0}) + 1
            ELSE YEAR({0})
        END""".format("{0}", Config.FISCAL_YEAR_START_MONTH)
    else:
        # For fiscal years starting Jan-Sep
        return """
        CASE 
            WHEN MONTH({0}) >= {1} THEN YEAR({0})
            ELSE YEAR({0}) - 1
        END""".format("{0}", Config.FISCAL_YEAR_START_MONTH)

def campus_filter(alias='p'):
    """Return SQL fragment for campus filtering. Empty string when All Campuses."""
    if Config.CAMPUS_ID > 0:
        if alias:
            return " AND {}.CampusId = {}".format(alias, Config.CAMPUS_ID)
        return " AND CampusId = {}".format(Config.CAMPUS_ID)
    return ""

def get_year_label():
    """Get label for year display"""
    if Config.USE_FISCAL_YEAR:
        return "FY"
    return ""

def print_report_styles():
    """Print report-specific CSS styles"""
    print("""
    <style>
    @media print {
        .btn { display: none; }
        .page-break { page-break-before: always; }
    }
    .table { table-layout: auto; width: 100%; }
    .table td, .table th { word-wrap: break-word; }
    canvas { max-width: 100%; }
    .panel {
        border: 1px solid #ddd;
        border-radius: 4px;
        margin-bottom: 20px;
    }
    .panel-body {
        padding: 15px;
    }
    .panel-heading {
        padding: 10px 15px;
        background-color: #f5f5f5;
        border-bottom: 1px solid #ddd;
        border-radius: 3px 3px 0 0;
    }
    .panel-primary .panel-body {
        background-color: #f0f4ff;
    }
    .panel-success .panel-body {
        background-color: #f0fff4;
    }
    .panel-info .panel-body {
        background-color: #f0faff;
    }
    .panel-danger .panel-body {
        background-color: #fff0f0;
    }
    h2 {
        margin-top: 30px;
        margin-bottom: 20px;
        border-bottom: 2px solid #eee;
        padding-bottom: 10px;
    }
    h3 {
        margin-top: 20px;
        margin-bottom: 15px;
    }
    </style>
    """)


def handle_baptism_drilldown():
    """Handle baptism drill-down AJAX requests"""
    import re as _re
    import urllib as _urllib
    _qs = str(model.QueryString) if hasattr(model, 'QueryString') else ''

    def _qs_param(name):
        _m = _re.search(name + r'=([^&]+)', _qs)
        if _m:
            return _urllib.unquote(_m.group(1))
        if hasattr(model.Data, name):
            return str(getattr(model.Data, name))
        return None

    _yr = _qs_param('yr')
    _bin = _qs_param('bin')

    print '<!-- AJAX_CONTENT_START -->'
    if _yr and _bin:
        _fiscal_year_sql = get_fiscal_year_sql().format("p.BaptismDate")

        _age_at_baptism = """CASE
            WHEN p.BDate IS NOT NULL THEN
                DATEDIFF(year, p.BDate, p.BaptismDate)
                - CASE WHEN DATEADD(year, DATEDIFF(year, p.BDate, p.BaptismDate), p.BDate) > p.BaptismDate THEN 1 ELSE 0 END
            ELSE p.Age
        END"""

        _age_filter = ''
        if _bin == 'Unknown':
            _age_filter = '({}) IS NULL'.format(_age_at_baptism)
        else:
            for _label, _min_age, _max_age in Config.BAPTISM_AGE_BINS:
                if _label == _bin:
                    if _max_age >= 150:
                        _age_filter = '({}) >= {}'.format(_age_at_baptism, _min_age)
                    else:
                        _age_filter = '({}) >= {} AND ({}) <= {}'.format(_age_at_baptism, _min_age, _age_at_baptism, _max_age)
                    break

        if _age_filter:
            _sql = """
            SELECT p.PeopleId, p.Name2, p.EmailAddress, p.Age, p.BaptismDate,
                   ({age_at_baptism}) AS AgeAtBaptism,
                   ms.Description AS MemberStatus
            FROM People p WITH (NOLOCK)
            LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
            WHERE p.BaptismDate IS NOT NULL
              AND p.IsDeceased = 0
              AND {fiscal_year} = {yr}
              AND {age_filter}{campus}
            ORDER BY p.BaptismDate DESC, p.Name2
            """.format(
                age_at_baptism=_age_at_baptism,
                fiscal_year=_fiscal_year_sql,
                yr=_yr,
                age_filter=_age_filter,
                campus=campus_filter()
            )
            _rows = list(q.QuerySql(_sql))

            if _rows:
                print '<table class="table table-striped table-condensed" style="margin:0">'
                print '<thead><tr><th>Name</th><th>Age at Baptism</th><th>Baptism Date</th><th>Status</th></tr></thead>'
                print '<tbody>'
                for _r in _rows:
                    _pid = safe_get_value(_r, 'PeopleId', '')
                    _name = safe_get_value(_r, 'Name2', '')
                    _aab = safe_get_value(_r, 'AgeAtBaptism', '')
                    _bd = safe_get_value(_r, 'BaptismDate', '')
                    _ms = safe_get_value(_r, 'MemberStatus', '')
                    _bd_str = ''
                    try:
                        _bd_str = _bd.ToString("MM/dd/yyyy") if _bd else ''
                    except:
                        _bd_str = str(_bd)[:10] if _bd else ''
                    print '<tr>'
                    print '<td><a href="/Person2/{}" target="_blank">{}</a></td>'.format(_pid, _name)
                    print '<td>{}</td>'.format(_aab if _aab != '' else '?')
                    print '<td>{}</td>'.format(_bd_str)
                    print '<td>{}</td>'.format(_ms)
                    print '</tr>'
                print '</tbody></table>'
            else:
                print '<p class="text-muted" style="padding:10px">No records found.</p>'
        else:
            print '<p class="text-muted" style="padding:10px">Invalid age bin.</p>'
    else:
        print '<p class="text-muted" style="padding:10px">Missing parameters.</p>'
    print '<!-- AJAX_CONTENT_END -->'


# ============================================================
# MAIN ENTRY POINT
# ============================================================
# Routing:
#   /PyScriptForm/  (GET)              -> Run the report (default)
#   /PyScriptForm/  (GET ?settings=1)  -> Config UI
#   /PyScriptForm/  (POST)             -> AJAX handlers
# ============================================================

# Check for baptism drill-down AJAX
_baptism_drilldown = False
try:
    if hasattr(Data, 'baptism_drill') and str(Data.baptism_drill) == 'true':
        _baptism_drilldown = True
except:
    pass

# Check for settings mode (GET ?settings=1)
_show_settings = False
try:
    if hasattr(Data, 'settings') and str(Data.settings) == '1':
        _show_settings = True
except:
    pass

# Check for campus override from URL (?campus=X)
try:
    if hasattr(Data, 'campus') and Data.campus is not None and str(Data.campus) != '':
        _active_config['campusId'] = int(str(Data.campus))
except:
    pass

if _baptism_drilldown:
    # Baptism age-bin drilldown (AJAX from report page).
    # The script is hit via /PyScriptForm/..., which renders model.Form
    # and IGNORES raw print output. Without this stdout-to-buffer wrap
    # the prints land in the void, the response body is empty (just the
    # TouchPoint chrome), the client can't find the AJAX_CONTENT markers,
    # and the modal shows "Could not parse response." Bug 2026-06-15.
    _old_stdout = sys.stdout
    _buffer = StringIO()
    sys.stdout = _buffer
    try:
        handle_baptism_drilldown()
    finally:
        sys.stdout = _old_stdout
    model.Form = _buffer.getvalue()
elif model.HttpMethod == "post":
    action = get_form_data('action', '')
    if action:
        # AJAX config handlers (save/load/campuses/programs)
        handle_ajax()
    else:
        print json.dumps({'success': False, 'message': 'No action specified'})
elif _show_settings:
    # GET with ?settings=1 -> Show config UI
    model.Header = TITLE + " - Settings"
    model.Form = generate_ui()
else:
    # GET (default) -> Run the report
    # Capture print output into model.Form so it renders inside the TP page frame
    _old_stdout = sys.stdout
    _buffer = StringIO()
    sys.stdout = _buffer
    try:
        print_report_styles()
        main()
    finally:
        sys.stdout = _old_stdout
    model.Form = _buffer.getvalue()

# End of report
