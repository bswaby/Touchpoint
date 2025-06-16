# ::START:: Program to Connect Group Analytics Dashboard - Working Version
# 
# Purpose: Analyze how effectively church programs drive attendance to Connect Groups
# Features: 
# - Shows percentage of program members who ACTUALLY attend Connect Groups (using Attend table)
# - Identifies members who started Connect Group attendance AFTER joining program
# - Calculates days between program enrollment and first Connect Group attendance
# - Analyzes individual program members (not families) for main metrics
# - Provides widget summary statistics
#
# Upload Instructions:
# 1. Click Admin > Advanced > Special Content > Python
# 2. Click New Python Script File  
# 3. Name the script "ProgramAttendanceAnalytics" and paste all this code
# 4. Test and optionally add to menu
#

# written by: Ben Swaby
# email: bswaby@fbchtn.org

# ::START:: Configuration and Imports
import traceback
model.Header = "Program to Connect Groups"

class ProgramAttendanceAnalyzer:
    def __init__(self):
        # ::STEP:: Define Program and Connect Group Configuration
        self.CHURCH_PROGRAM_ID = 1128  # Connect Group ProgId - the target we want people to join
        self.PROGRAM_IDS = {
            1108: "FMC",                    
            1158: "American Heritage Girls", 
            1116: "MOPS",                   
            1143: "Peer Place",
            1149: "Rooted Enrichment",
            1123: "VBS K-5",
            1135: "VBS Preschool",
            1109: "WEE"
        }

    def get_organization_ids_for_program(self, prog_id):
        """
        ::START:: Get OrganizationIds for a Program (Working Version)
        Uses QuerySqlInt for counts and careful result handling for data
        """
        try:
            # ::STEP:: Check if ProgId exists using QuerySqlInt (we know this works)
            count_sql = "SELECT COUNT(*) FROM OrganizationStructure WHERE ProgId = {}".format(prog_id)
            prog_count = q.QuerySqlInt(count_sql)
            print "<!-- DEBUG: ProgId {} found in {} OrganizationStructure records -->".format(prog_id, prog_count)
            
            if prog_count == 0:
                print "<!-- DEBUG: ProgId {} not found in OrganizationStructure table -->".format(prog_id)
                return [], []
            
            # ::STEP:: Get Active OrganizationIds for this Program using QuerySql
            org_sql = """
            SELECT OrgId, Organization, OrgStatus
            FROM OrganizationStructure 
            WHERE ProgId = {}
            AND OrgStatus = 'Active'
            ORDER BY OrgId
            """.format(prog_id)
            
            result = q.QuerySql(org_sql)
            if result and len(result) > 0:
                org_ids = []
                org_names = []
                
                print "<!-- DEBUG: ProgId {} returned {} rows from QuerySql -->".format(prog_id, len(result))
                
                # ::STEP:: Process each row carefully
                for i, row in enumerate(result):
                    try:
                        # In TouchPoint Python, SQL results often have Column1, Column2, Column3, etc.
                        # Let's try both named access and positional access
                        org_id = None
                        org_name = "Unknown"
                        org_status = "Unknown"
                        
                        # Try named attribute access first
                        try:
                            if hasattr(row, 'OrgId'):
                                org_id = int(row.OrgId)
                            if hasattr(row, 'Organization'):
                                org_name = str(row.Organization)
                            if hasattr(row, 'OrgStatus'):
                                org_status = str(row.OrgStatus)
                        except:
                            pass
                        
                        # If named access didn't work, try Column1, Column2, Column3
                        if org_id is None:
                            try:
                                if hasattr(row, 'Column1'):
                                    org_id = int(row.Column1)
                                if hasattr(row, 'Column2'):
                                    org_name = str(row.Column2)
                                if hasattr(row, 'Column3'):
                                    org_status = str(row.Column3)
                            except:
                                pass
                        
                        # If we got valid data, add it
                        if org_id is not None:
                            org_ids.append(org_id)
                            org_names.append(org_name)
                            print "<!-- DEBUG: Added OrgId {} ({}) Status: {} -->".format(org_id, org_name, org_status)
                        else:
                            print "<!-- DEBUG: Could not parse row {} for ProgId {} -->".format(i, prog_id)
                            
                    except Exception as e:
                        print "<!-- DEBUG: Error processing row {} for ProgId {}: {} -->".format(i, prog_id, str(e))
                        continue
                
                print "<!-- DEBUG: ProgId {} resolved to {} active organizations: {} -->".format(prog_id, len(org_ids), org_ids)
                return org_ids, org_names
            else:
                print "<!-- DEBUG: ProgId {} QuerySql returned no results -->".format(prog_id)
                return [], []
                
        except Exception as e:
            print "<!-- ERROR: Getting org IDs for program {}: {} -->".format(prog_id, str(e))
            return [], []

    def analyze_program_simple(self, prog_id):
        """
        ::START:: Simplified Program Analysis (Working Version)
        Uses QuerySqlInt for scalar values and careful handling for complex queries
        """
        try:
            # ::STEP:: Get All OrganizationIds for this Program
            org_ids, org_names = self.get_organization_ids_for_program(prog_id)
            if not org_ids:
                print "<!-- No organizations found for program {} -->".format(prog_id)
                return {
                    'total_members': 0,
                    'connect_group_attenders': 0, 
                    'converted_count': 0,
                    'conversion_rate': 0.0,
                    'attendance_rate': 0.0,
                    'avg_days_to_conversion': 0.0,
                    'organization_count': 0,
                    'organization_names': []
                }
            
            # ::STEP:: Get Connect Group OrganizationIds 
            connect_group_org_ids, _ = self.get_organization_ids_for_program(self.CHURCH_PROGRAM_ID)
            if not connect_group_org_ids:
                # Fallback: assume CHURCH_PROGRAM_ID is already an OrganizationId
                connect_group_org_ids = [self.CHURCH_PROGRAM_ID]
            
            # ::STEP:: Convert lists to comma-separated strings for SQL IN clauses
            org_ids_str = ",".join(str(oid) for oid in org_ids)
            connect_group_org_ids_str = ",".join(str(cid) for cid in connect_group_org_ids)
            
            print "<!-- Program {} has {} organizations: {} -->".format(prog_id, len(org_ids), org_names[:3])
            print "<!-- Connect Groups program {} resolves to {} organizations -->".format(self.CHURCH_PROGRAM_ID, len(connect_group_org_ids))
            
            # ::STEP:: Get Program Members Count using QuerySqlInt with DISTINCT to avoid double-counting
            member_sql = """
            SELECT COUNT(DISTINCT om.PeopleId) 
            FROM OrganizationMembers om
            WHERE om.OrganizationId IN ({}) 
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
            """.format(org_ids_str)
            
            total_members = q.QuerySqlInt(member_sql)
            print "<!-- DEBUG: ProgId {} has {} distinct people connected -->".format(prog_id, total_members)
            
            # ::STEP:: Get Program Members Who Actually Attend Connect Groups
            connect_group_attender_sql = """
            SELECT COUNT(DISTINCT pm.PeopleId)
            FROM OrganizationMembers pm
            INNER JOIN Attend a ON pm.PeopleId = a.PeopleId
            WHERE pm.OrganizationId IN ({})
            AND a.OrganizationId IN ({})
            AND a.AttendanceFlag = 1
            AND a.MeetingDate >= DATEADD(month, -12, GETDATE())
            AND (pm.InactiveDate IS NULL OR pm.InactiveDate > GETDATE())
            """.format(org_ids_str, connect_group_org_ids_str)
            
            connect_group_attenders = q.QuerySqlInt(connect_group_attender_sql)
            print "<!-- DEBUG: ProgId {} has {} connect group attenders -->".format(prog_id, connect_group_attenders)
            
            # ::STEP:: Get Conversion Count (People who attended Connect Groups AFTER joining program)
            conversion_sql = """
            SELECT COUNT(DISTINCT pm.PeopleId)
            FROM OrganizationMembers pm
            INNER JOIN (
                SELECT PeopleId, MIN(MeetingDate) as FirstConnectGroupAttendance
                FROM Attend 
                WHERE OrganizationId IN ({})
                AND AttendanceFlag = 1
                GROUP BY PeopleId
            ) first_connect_group ON pm.PeopleId = first_connect_group.PeopleId
            WHERE pm.OrganizationId IN ({})
            AND pm.EnrollmentDate < first_connect_group.FirstConnectGroupAttendance
            AND (pm.InactiveDate IS NULL OR pm.InactiveDate > GETDATE())
            """.format(connect_group_org_ids_str, org_ids_str)
            
            converted_count = q.QuerySqlInt(conversion_sql)
            print "<!-- DEBUG: ProgId {} has {} conversions -->".format(prog_id, converted_count)
            
            # ::STEP:: Calculate Rates
            attendance_rate = (float(connect_group_attenders) / float(total_members) * 100.0) if total_members > 0 else 0.0
            conversion_rate = (float(converted_count) / float(total_members) * 100.0) if total_members > 0 else 0.0
            
            return {
                'total_members': total_members,
                'connect_group_attenders': connect_group_attenders,
                'converted_count': converted_count,
                'conversion_rate': conversion_rate,
                'attendance_rate': attendance_rate,
                'avg_days_to_conversion': 0.0,  # We'll calculate this later if needed
                'organization_count': len(org_ids),
                'organization_names': org_names
            }
            
        except Exception as e:
            print "<!-- ERROR: Analyzing program {}: {} -->".format(prog_id, str(e))
            return {
                'total_members': 0,
                'connect_group_attenders': 0, 
                'converted_count': 0,
                'conversion_rate': 0.0,
                'attendance_rate': 0.0,
                'avg_days_to_conversion': 0.0,
                'organization_count': 0,
                'organization_names': []
            }

    def get_widget_stats(self):
        """
        ::START:: Widget Statistics Generation
        Returns summary stats for dashboard widget display
        """
        try:
            # ::STEP:: Calculate Overall Program Effectiveness Using DISTINCT People
            # Get all active organization IDs for our programs
            all_program_org_ids = []
            for program_id in self.PROGRAM_IDS:
                org_ids, _ = self.get_organization_ids_for_program(program_id)
                all_program_org_ids.extend(org_ids)
            
            if all_program_org_ids:
                all_org_ids_str = ",".join(str(oid) for oid in all_program_org_ids)
                
                # Get DISTINCT people across all programs
                distinct_members_sql = """
                SELECT COUNT(DISTINCT om.PeopleId)
                FROM OrganizationMembers om
                WHERE om.OrganizationId IN ({})
                AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
                """.format(all_org_ids_str)
                
                total_program_members = q.QuerySqlInt(distinct_members_sql)
                
                # Get Connect Group org IDs
                connect_group_org_ids, _ = self.get_organization_ids_for_program(self.CHURCH_PROGRAM_ID)
                if not connect_group_org_ids:
                    connect_group_org_ids = [self.CHURCH_PROGRAM_ID]
                
                connect_group_org_ids_str = ",".join(str(cid) for cid in connect_group_org_ids)
                
                # Get DISTINCT attenders
                distinct_attenders_sql = """
                SELECT COUNT(DISTINCT pm.PeopleId)
                FROM OrganizationMembers pm
                INNER JOIN Attend a ON pm.PeopleId = a.PeopleId
                WHERE pm.OrganizationId IN ({})
                AND a.OrganizationId IN ({})
                AND a.AttendanceFlag = 1
                AND a.MeetingDate >= DATEADD(month, -12, GETDATE())
                AND (pm.InactiveDate IS NULL OR pm.InactiveDate > GETDATE())
                """.format(all_org_ids_str, connect_group_org_ids_str)
                
                total_connect_group_attenders = q.QuerySqlInt(distinct_attenders_sql)
                
                # Get DISTINCT conversions
                distinct_conversions_sql = """
                SELECT COUNT(DISTINCT pm.PeopleId)
                FROM OrganizationMembers pm
                INNER JOIN (
                    SELECT PeopleId, MIN(MeetingDate) as FirstConnectGroupAttendance
                    FROM Attend 
                    WHERE OrganizationId IN ({})
                    AND AttendanceFlag = 1
                    GROUP BY PeopleId
                ) first_connect_group ON pm.PeopleId = first_connect_group.PeopleId
                WHERE pm.OrganizationId IN ({})
                AND pm.EnrollmentDate < first_connect_group.FirstConnectGroupAttendance
                AND (pm.InactiveDate IS NULL OR pm.InactiveDate > GETDATE())
                """.format(connect_group_org_ids_str, all_org_ids_str)
                
                total_converted = q.QuerySqlInt(distinct_conversions_sql)
                
            else:
                total_program_members = 0
                total_connect_group_attenders = 0
                total_converted = 0
            
            if total_program_members > 0:
                connect_group_attendance_rate = (float(total_connect_group_attenders) / float(total_program_members)) * 100.0
                conversion_rate = (float(total_converted) / float(total_program_members)) * 100.0
            else:
                connect_group_attendance_rate = 0.0
                conversion_rate = 0.0
                
            widget_html = """
            <div class="pta-widget-container">
                <h2 class="pta-widget-title">Program Impact</h2>
                <div class="pta-widget-number">{:.1f}%</div>
                <div class="pta-widget-subtitle">of unique program participants attend Connect Groups</div>
                <div class="pta-widget-converts">{} New Connects</div>
                <div class="pta-widget-footer">Programs → Connect Groups Pipeline</div>
            </div>
            <style>
            .pta-widget-container {{ 
                text-align: center; 
                padding: 20px; 
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                color: white; 
                border-radius: 10px; 
                margin: 10px; 
                font-family: Arial, sans-serif;
            }}
            .pta-widget-title {{ 
                margin: 0; 
                font-size: 24px; 
                font-weight: bold;
            }}
            .pta-widget-number {{ 
                font-size: 36px; 
                font-weight: bold; 
                margin: 10px 0; 
            }}
            .pta-widget-subtitle {{ 
                font-size: 14px; 
            }}
            .pta-widget-converts {{ 
                font-size: 18px; 
                margin-top: 10px; 
                color: #ffeb3b; 
                font-weight: bold;
            }}
            .pta-widget-footer {{ 
                font-size: 12px; 
                margin-top: 5px; 
                opacity: 0.9;
            }}
            </style>
            """.format(connect_group_attendance_rate, total_converted)
            
            return widget_html
            
        except Exception as e:
            return "<div style='color: red; padding: 20px;'>Widget Error: {}</div>".format(str(e))

    def generate_dashboard_html(self):
        """
        ::START:: Main Dashboard HTML Generation
        Creates the complete dashboard interface
        """
        css = """
        <style>
        .pta-dashboard-container { 
            max-width: 1200px; 
            margin: 0 auto; 
            font-family: Arial, sans-serif; 
            padding: 20px;
        }
        .pta-program-card { 
            background: white; 
            border-radius: 8px; 
            padding: 20px; 
            margin: 15px 0; 
            box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
            border-left: 5px solid #3498db;
        }
        .pta-program-header { 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            margin-bottom: 15px; 
            flex-wrap: wrap;
        }
        .pta-program-title { 
            font-size: 24px; 
            font-weight: bold; 
            color: #2c3e50; 
        }
        .pta-stats-grid { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); 
            gap: 15px; 
            margin: 20px 0; 
        }
        .pta-stat-box { 
            background: #f8f9fa; 
            padding: 15px; 
            border-radius: 5px; 
            text-align: center; 
            border: 1px solid #e9ecef;
        }
        .pta-stat-number { 
            font-size: 28px; 
            font-weight: bold; 
            color: #3498db; 
            margin-bottom: 5px;
        }
        .pta-stat-label { 
            font-size: 14px; 
            color: #666; 
        }
        .pta-loading { 
            text-align: center; 
            padding: 50px; 
            font-size: 18px; 
            color: #666; 
        }
        .pta-success-highlight { 
            background: #d4edda; 
            border-left-color: #28a745 !important; 
        }
        .pta-warning-highlight { 
            background: #fff3cd; 
            border-left-color: #ffc107 !important; 
        }
        .pta-header-section {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            text-align: center;
        }
        .pta-summary-stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        .pta-summary-stat {
            background: rgba(255,255,255,0.2);
            padding: 20px;
            border-radius: 8px;
            text-align: center;
        }
        .pta-summary-number {
            font-size: 36px;
            font-weight: bold;
            margin-bottom: 5px;
        }
        .pta-summary-label {
            font-size: 14px;
            opacity: 0.9;
        }
        @media (max-width: 768px) {
            .pta-program-header {
                flex-direction: column;
                align-items: flex-start;
            }
            .pta-stats-grid {
                grid-template-columns: 1fr 1fr;
            }
            .pta-summary-stats {
                grid-template-columns: 1fr 1fr;
            }
        }
        </style>
        """
        
        html = css + """
        <div class="pta-dashboard-container">
            <div class="pta-header-section">
                <h1 style="margin: 0 0 20px 0; font-size: 32px;">Program to Connect Group Analytics<svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
                    <!-- Text portion - TP -->
                    <text x="100" y="120" font-family="Arial, sans-serif" font-weight="bold" font-size="60" fill="#333333">TP</text>
                    
                    <!-- Circular element -->
                    <g transform="translate(190, 107)">
                      <!-- Outer circle -->
                      <circle cx="0" cy="0" r="13.5" fill="#0099FF"/>
                      
                      <!-- White middle circle -->
                      <circle cx="0" cy="0" r="10.5" fill="white"/>
                      
                      <!-- Inner circle -->
                      <circle cx="0" cy="0" r="7.5" fill="#0099FF"/>
                      
                      <!-- X crossing through the circles -->
                      <path d="M-9 -9 L9 9 M-9 9 L9 -9" stroke="white" stroke-width="1.8" stroke-linecap="round"/>
                    </g>
                    
                    <!-- Single "i" letter to the right -->
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">si</text>
                  </svg></h1>
                <p style="margin: 0; font-size: 16px; opacity: 0.9;">
                    Tracking how effectively church programs drive attendance to Connect Groups
                </p>
                <div style="margin-top: 15px; font-size: 14px; opacity: 0.8; font-style: italic;">
                    📊 Overall summary shows unique people across all programs (no double-counting)<br>
                    📈 Individual program metrics count people separately per program<br>
                    💡 Analysis based on Connect Group attendance in the last 12 months
                </div>
            </div>
            
            <div id="pta-loading" class="pta-loading">
                <div>📊 Analyzing program effectiveness...</div>
                <div style="margin-top: 10px;">This may take a few moments</div>
            </div>
            
            <div id="pta-dashboard-content" style="display: none;">
        """
        
        try:
            # ::STEP:: Calculate Summary Statistics Using DISTINCT People Across All Programs
            program_count = len(self.PROGRAM_IDS)
            
            # Get all active organization IDs for our programs
            all_program_org_ids = []
            program_stats = {}
            for program_id in self.PROGRAM_IDS:
                stats = self.analyze_program_simple(program_id)
                program_stats[program_id] = stats
                org_ids, _ = self.get_organization_ids_for_program(program_id)
                all_program_org_ids.extend(org_ids)
            
            # Calculate DISTINCT totals across all programs for the summary
            if all_program_org_ids:
                all_org_ids_str = ",".join(str(oid) for oid in all_program_org_ids)
                
                # Get DISTINCT people across all programs (avoids double-counting)
                distinct_members_sql = """
                SELECT COUNT(DISTINCT om.PeopleId)
                FROM OrganizationMembers om
                WHERE om.OrganizationId IN ({})
                AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
                """.format(all_org_ids_str)
                
                total_program_members = q.QuerySqlInt(distinct_members_sql)
                
                # Get Connect Group org IDs
                connect_group_org_ids, _ = self.get_organization_ids_for_program(self.CHURCH_PROGRAM_ID)
                if not connect_group_org_ids:
                    connect_group_org_ids = [self.CHURCH_PROGRAM_ID]
                
                connect_group_org_ids_str = ",".join(str(cid) for cid in connect_group_org_ids)
                
                # Get DISTINCT people who attend both programs and Connect Groups
                distinct_attenders_sql = """
                SELECT COUNT(DISTINCT pm.PeopleId)
                FROM OrganizationMembers pm
                INNER JOIN Attend a ON pm.PeopleId = a.PeopleId
                WHERE pm.OrganizationId IN ({})
                AND a.OrganizationId IN ({})
                AND a.AttendanceFlag = 1
                AND a.MeetingDate >= DATEADD(month, -12, GETDATE())
                AND (pm.InactiveDate IS NULL OR pm.InactiveDate > GETDATE())
                """.format(all_org_ids_str, connect_group_org_ids_str)
                
                total_connect_group_attenders = q.QuerySqlInt(distinct_attenders_sql)
                
                # Get DISTINCT conversions
                distinct_conversions_sql = """
                SELECT COUNT(DISTINCT pm.PeopleId)
                FROM OrganizationMembers pm
                INNER JOIN (
                    SELECT PeopleId, MIN(MeetingDate) as FirstConnectGroupAttendance
                    FROM Attend 
                    WHERE OrganizationId IN ({})
                    AND AttendanceFlag = 1
                    GROUP BY PeopleId
                ) first_connect_group ON pm.PeopleId = first_connect_group.PeopleId
                WHERE pm.OrganizationId IN ({})
                AND pm.EnrollmentDate < first_connect_group.FirstConnectGroupAttendance
                AND (pm.InactiveDate IS NULL OR pm.InactiveDate > GETDATE())
                """.format(connect_group_org_ids_str, all_org_ids_str)
                
                total_converted = q.QuerySqlInt(distinct_conversions_sql)
                
            else:
                total_program_members = 0
                total_connect_group_attenders = 0
                total_converted = 0
            
            overall_attendance_rate = (float(total_connect_group_attenders) / float(total_program_members) * 100.0) if total_program_members > 0 else 0.0
            overall_conversion_rate = (float(total_converted) / float(total_program_members) * 100.0) if total_program_members > 0 else 0.0
            
            # ::STEP:: Add Summary Section
            html += """
                <div class="pta-program-card">
                    <div class="pta-program-header">
                        <div class="pta-program-title">📈 Overall Program Impact Summary</div>
                        <div style="font-size: 14px; color: #666; margin-top: 10px;">
                            Our total reach: unique people connected across all programs
                        </div>
                    </div>
                    <div class="pta-summary-stats">
                        <div class="pta-summary-stat">
                            <div class="pta-summary-number">{}</div>
                            <div class="pta-summary-label">Programs Analyzed</div>
                        </div>
                        <div class="pta-summary-stat">
                            <div class="pta-summary-number">{}</div>
                            <div class="pta-summary-label">Unique People Reached</div>
                        </div>
                        <div class="pta-summary-stat">
                            <div class="pta-summary-number">{:.1f}%</div>
                            <div class="pta-summary-label">Also Attend Connect Groups</div>
                        </div>
                        <div class="pta-summary-stat">
                            <div class="pta-summary-number">{}</div>
                            <div class="pta-summary-label">Started Connect Groups After Program</div>
                        </div>
                        <div class="pta-summary-stat">
                            <div class="pta-summary-number">{:.1f}%</div>
                            <div class="pta-summary-label">Program → Connect Group Success Rate</div>
                        </div>
                    </div>
                    <div style="margin-top: 15px; padding: 15px; background: rgba(255,255,255,0.1); border-radius: 5px; font-size: 14px;">
                        <strong>What these numbers mean:</strong><br>
                        • <strong>Unique People Reached:</strong> Distinct individuals connected to any of our programs (people in multiple programs counted once)<br>
                        • <strong>Also Attend Connect Groups:</strong> Unique people who participate in both programs and Connect Groups in the last 12 months<br>
                        • <strong>Started Connect Groups After Program:</strong> People who joined a Connect Group AFTER connecting to any program (true conversions)<br>
                        • <strong>Success Rate:</strong> Percentage of unique people successfully connected to ongoing community<br>
                        <em>Note: Overall summary counts each person once. Individual program cards below count people separately per program.</em>
                    </div>
                </div>
            """.format(
                program_count,
                total_program_members,
                overall_attendance_rate,
                total_converted,
                overall_conversion_rate
            )
            
            # ::STEP:: Generate Individual Program Analysis
            for program_id, program_name in sorted(self.PROGRAM_IDS.items(), key=lambda x: x[1]):
                stats = program_stats[program_id]
                
                card_class = "pta-program-card"
                if stats['conversion_rate'] > 20:
                    card_class += " pta-success-highlight"
                elif stats['conversion_rate'] > 10:
                    card_class += " pta-warning-highlight"
                
                html += """
                    <div class="{}">
                        <div class="pta-program-header">
                            <div class="pta-program-title">{} ({})</div>
                            <div style="font-size: 14px; color: #666;">
                                {} people connected across {} sub-organizations
                            </div>
                            <div style="font-size: 12px; color: #888; margin-top: 5px;">
                                Sub-organizations: {}
                            </div>
                        </div>
                        
                        <div class="pta-stats-grid">
                            <div class="pta-stat-box">
                                <div class="pta-stat-number">{:.1f}%</div>
                                <div class="pta-stat-label">Also in Connect Groups</div>
                            </div>
                            <div class="pta-stat-box">
                                <div class="pta-stat-number">{}</div>
                                <div class="pta-stat-label">Connect Group Members</div>
                            </div>
                            <div class="pta-stat-box">
                                <div class="pta-stat-number">{}</div>
                                <div class="pta-stat-label">Program → Connect Group</div>
                            </div>
                            <div class="pta-stat-box">
                                <div class="pta-stat-number">{:.1f}%</div>
                                <div class="pta-stat-label">Connection Success Rate</div>
                            </div>
                        </div>
                        
                        <div style="margin-top: 15px; padding: 10px; background: #f8f9fa; border-radius: 5px; font-size: 13px; color: #666;">
                            <strong>What this shows:</strong> Out of {} people directly connected to the {} program, {} ({}%) also attend Connect Groups. 
                            {} people joined Connect Groups AFTER connecting to this program, showing a {}% success rate at connecting people to ongoing community.
                            <div style="margin-top: 8px; font-style: italic;">
                                <strong>Note:</strong> This counts people connected to this specific program. People in multiple programs are counted in each program's metrics.
                            </div>
                        </div>
                    </div>
                """.format(
                    card_class,
                    program_name,
                    program_id,
                    stats['total_members'],
                    stats['organization_count'],
                    ", ".join(stats['organization_names'][:3]) + ("..." if len(stats['organization_names']) > 3 else ""),
                    stats['attendance_rate'],
                    stats['connect_group_attenders'],
                    stats['converted_count'],
                    stats['conversion_rate'],
                    # Explanation sentence
                    stats['total_members'],
                    program_name,
                    stats['connect_group_attenders'],
                    stats['attendance_rate'],
                    stats['converted_count'],
                    stats['conversion_rate']
                )
            
        except Exception as e:
            html += """
                <div class="pta-program-card" style="border-left-color: #e74c3c; background: #fdebea;">
                    <h3 style="color: #c0392b;">⚠️ Analysis Error</h3>
                    <p>An error occurred while analyzing programs: {}</p>
                    <pre style="background: #f8f8f8; padding: 15px; border-radius: 5px; font-size: 12px; overflow: auto;">
{}
                    </pre>
                </div>
            """.format(str(e), traceback.format_exc())
        
        html += """
            </div>
        </div>
        
        <script>
            setTimeout(function() {
                var loading = document.getElementById('pta-loading');
                var content = document.getElementById('pta-dashboard-content');
                if (loading) loading.style.display = 'none';
                if (content) content.style.display = 'block';
            }, 1000);
        </script>
        """
        
        return html

# ::START:: Main Controller Logic
def main():
    """
    ::START:: Main Application Controller
    Handles routing and request processing
    """
    try:
        analyzer = ProgramAttendanceAnalyzer()
        
        # ::STEP:: Check for Widget Request
        widget_request = False
        try:
            if hasattr(model.Data, 'widgetstat') and str(model.Data.widgetstat) == '1':
                widget_request = True
        except:
            pass
        
        # ::STEP:: Route Request
        if widget_request:
            print analyzer.get_widget_stats()
        else:
            print analyzer.generate_dashboard_html()
            
    except Exception as e:
        print """
        <div style="max-width: 800px; margin: 50px auto; padding: 30px; background: #fdebea; border-radius: 10px; border-left: 5px solid #e74c3c; font-family: Arial, sans-serif;">
            <h2 style="color: #c0392b; margin: 0 0 20px 0;">🚨 Dashboard Error</h2>
            <p>An error occurred: {}</p>
            <pre style="background: #f8f8f8; padding: 15px; border-radius: 5px; font-size: 12px; overflow: auto;">
{}
            </pre>
        </div>
        """.format(str(e), traceback.format_exc())

# ::START:: Application Entry Point
# Execute main controller
main()
