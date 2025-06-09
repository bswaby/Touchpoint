# TouchPoint Unified Security Monitoring System - All-in-One Dashboard
# 
# Purpose: Single script that provides KPI overview and access to multiple security dashboards
# Features: Real-time KPI summary, integrated dashboard selector, and multiple monitoring tools
# 
# --Upload Instructions Start--
# To upload code to Touchpoint, use the following steps:
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File  
# 3. Name the Python script "UnifiedSecurityMonitoring" and paste all this code
# 4. Test and optionally add to menu
# --Upload Instructions End--
#
#written by: Ben Swaby
#email: bswaby@fbchtn.org

import traceback
from datetime import datetime, timedelta

# ::START:: Main Controller
def main():
    """Main function to handle unified security system routing"""
    try:
        # Get view parameter and action
        view = None
        action = None
        
        try:
            if hasattr(model, 'Data'):
                view = str(getattr(model.Data, 'view', ''))
                action = str(getattr(model.Data, 'action', ''))
        except:
            pass
        
        # Route to appropriate view
        if view == 'highrisk':
            show_back_to_hub_button()
            monitor = HighRiskUserMonitor()
            
            if action == 'generate_highrisk_report':
                params = get_highrisk_params()
                monitor.generate_highrisk_report(params)
            else:
                monitor.get_configuration_form()
        
        elif view == 'enhanced':
            show_back_to_hub_button()
            dashboard = EnhancedSecurityDashboard()
            
            if action == 'generate_report':
                params = get_enhanced_params()
                dashboard.generate_security_report(params)
            else:
                dashboard.get_configuration_form()
                
        else:
            # Default - Show KPI overview
            hub = UnifiedSecurityHub()
            hub.show_kpi_overview()
            
    except Exception as e:
        print_critical_error("Main Controller", e)

# ::START:: Unified Security Hub
class UnifiedSecurityHub:
    """Main hub for unified security monitoring system"""
    
    def __init__(self):
        self.kpi_lookback_days = 7
        self.critical_threshold = 20
        self.high_risk_threshold = 15
    
    def show_kpi_overview(self):
        """Display KPI overview and dashboard selector"""
        try:
            # ::STEP:: Fetch 7-day KPI Data
            sql_kpi = """
                DECLARE @StartDate DATETIME = DATEADD(DAY, -7, GETDATE());
                
                SELECT 
                    -- Total Failed Attempts
                    COUNT(*) AS TotalFailedAttempts,
                    COUNT(DISTINCT ClientIp) AS UniqueThreats,
                    COUNT(DISTINCT UserId) AS UsersAffected,
                    COUNT(DISTINCT CAST(ActivityDate AS DATE)) AS DaysWithActivity,
                    
                    -- Critical Threats (20+ attempts)
                    (SELECT COUNT(DISTINCT ClientIp) 
                     FROM ActivityLog 
                     WHERE Activity LIKE '%failed%' 
                     AND ActivityDate >= @StartDate
                     GROUP BY ClientIp 
                     HAVING COUNT(*) >= 20) AS CriticalThreats,
                    
                    -- Account Lockouts
                    SUM(CASE WHEN Activity LIKE '%locked%' THEN 1 ELSE 0 END) AS AccountLockouts,
                    
                    -- Password Resets
                    SUM(CASE WHEN Activity LIKE '%ForgotPassword%' THEN 1 ELSE 0 END) AS PasswordResets,
                    
                    -- Latest Activity
                    MAX(ActivityDate) AS LatestActivity,
                    
                    -- High Risk Users
                    (SELECT COUNT(DISTINCT u.UserId)
                     FROM ActivityLog al
                     JOIN Users u ON al.UserId = u.UserId
                     WHERE al.Activity LIKE '%failed%'
                     AND al.ActivityDate >= @StartDate
                     GROUP BY u.UserId
                     HAVING COUNT(*) >= 15) AS HighRiskUsers
                     
                FROM ActivityLog
                WHERE Activity LIKE '%failed%'
                AND ActivityDate >= @StartDate;
            """
            
            # Get critical IPs for alert display
            sql_critical_ips = """
                SELECT TOP 5
                    ClientIp,
                    COUNT(*) AS Attempts,
                    COUNT(DISTINCT UserId) AS UsersTargeted,
                    MAX(ActivityDate) AS LastSeen
                FROM ActivityLog
                WHERE Activity LIKE '%failed%'
                AND ActivityDate >= DATEADD(DAY, -7, GETDATE())
                GROUP BY ClientIp
                HAVING COUNT(*) >= 20
                ORDER BY Attempts DESC;
            """
            
            kpi_data = q.QuerySql(sql_kpi)
            critical_ips = q.QuerySql(sql_critical_ips)
            
            # ::STEP:: Generate KPI Overview Page
            print """
            <style>
            .unifsec2025_container {
                max-width: 1600px;
                margin: 0 auto;
                padding: 20px;
                font-family: Arial, sans-serif;
            }
            .unifsec2025_header {
                background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
                color: white;
                padding: 30px;
                border-radius: 10px;
                margin-bottom: 30px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
                text-align: center;
            }
            .unifsec2025_header h1 {
                margin: 0 0 10px 0;
                font-size: 32px;
            }
            .unifsec2025_header p {
                margin: 0;
                opacity: 0.9;
                font-size: 16px;
            }
            .unifsec2025_kpi_grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                gap: 20px;
                margin-bottom: 30px;
            }
            .unifsec2025_kpi_card {
                background: white;
                border-radius: 8px;
                padding: 20px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                border-left: 4px solid #007bff;
                transition: transform 0.2s ease;
                text-align: center;
            }
            .unifsec2025_kpi_card:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            }
            .unifsec2025_kpi_card.critical {
                border-left-color: #dc3545;
                background: #fff5f5;
            }
            .unifsec2025_kpi_card.warning {
                border-left-color: #ffc107;
                background: #fffaf0;
            }
            .unifsec2025_kpi_card.success {
                border-left-color: #28a745;
                background: #f0fff4;
            }
            .unifsec2025_kpi_value {
                font-size: 36px;
                font-weight: bold;
                color: #212529;
                margin: 10px 0;
            }
            .unifsec2025_kpi_label {
                font-size: 14px;
                color: #6c757d;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            .unifsec2025_kpi_subtitle {
                font-size: 12px;
                color: #6c757d;
                margin-top: 5px;
            }
            .unifsec2025_critical_alert {
                background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
                border: 2px solid #dc3545;
                border-radius: 8px;
                padding: 20px;
                margin-bottom: 30px;
                color: #721c24;
            }
            .unifsec2025_dashboard_selector {
                background: white;
                border-radius: 10px;
                padding: 30px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                margin-bottom: 30px;
            }
            .unifsec2025_dashboard_grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
                gap: 20px;
                margin-top: 20px;
            }
            .unifsec2025_dashboard_card {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                padding: 25px;
                text-align: center;
                transition: all 0.3s ease;
                background: white;
            }
            .unifsec2025_dashboard_card:hover {
                border-color: #007bff;
                background: #f8f9fa;
                transform: translateY(-3px);
                box-shadow: 0 4px 12px rgba(0,123,255,0.2);
            }
            .unifsec2025_dashboard_form {
                margin-top: 15px;
            }
            .unifsec2025_dashboard_btn {
                display: inline-block;
                padding: 12px 24px;
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 6px;
                font-weight: bold;
                font-size: 16px;
                cursor: pointer;
                transition: background-color 0.3s ease;
            }
            .unifsec2025_dashboard_btn:hover {
                background-color: #0056b3;
            }
            .unifsec2025_dashboard_btn.enhanced {
                background-color: #6f42c1;
            }
            .unifsec2025_dashboard_btn.enhanced:hover {
                background-color: #5a359a;
            }
            .unifsec2025_dashboard_btn.highrisk {
                background-color: #dc3545;
            }
            .unifsec2025_dashboard_btn.highrisk:hover {
                background-color: #c82333;
            }
            .unifsec2025_dashboard_btn.enhanced {
                background-color: #007bff;
            }
            .unifsec2025_dashboard_btn.enhanced:hover {
                background-color: #0056b3;
            }
            .unifsec2025_quick_actions {
                background: #f8f9fa;
                border-radius: 8px;
                padding: 20px;
                margin-bottom: 30px;
                text-align: center;
            }
            .unifsec2025_action_btn {
                display: inline-block;
                padding: 8px 16px;
                margin: 5px;
                background-color: #6c757d;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                cursor: pointer;
                transition: background-color 0.3s ease;
            }
            .unifsec2025_action_btn:hover {
                background-color: #5a6268;
            }
            .unifsec2025_action_btn.emergency {
                background-color: #dc3545;
            }
            .unifsec2025_action_btn.emergency:hover {
                background-color: #c82333;
            }
            </style>
            
            <div class="unifsec2025_container">
                <!-- Header -->
                <div class="unifsec2025_header">
                    <h1><i class="fa fa-shield"></i> Account Security Monitoring <svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
                    <p>Real-time security overview and threat analysis for the last 7 days</p>
                </div>
            """
            
            # Process KPI data
            if kpi_data and len(kpi_data) > 0:
                kpi = list(kpi_data)[0] if hasattr(kpi_data, '__iter__') else kpi_data
                
                total_attempts = getattr(kpi, 'TotalFailedAttempts', 0)
                unique_threats = getattr(kpi, 'UniqueThreats', 0)
                users_affected = getattr(kpi, 'UsersAffected', 0)
                days_active = getattr(kpi, 'DaysWithActivity', 0)
                critical_threats = getattr(kpi, 'CriticalThreats', 0) or 0
                lockouts = getattr(kpi, 'AccountLockouts', 0)
                password_resets = getattr(kpi, 'PasswordResets', 0)
                high_risk_users = getattr(kpi, 'HighRiskUsers', 0) or 0
                latest_activity = getattr(kpi, 'LatestActivity', None)
                
                # Display critical alert if needed
                if critical_threats > 0 and critical_ips:
                    print """
                    <div class="unifsec2025_critical_alert">
                        <h3 style="margin-top: 0;"><i class="fa fa-exclamation-triangle"></i> CRITICAL SECURITY ALERT</h3>
                        <p><strong>{} critical threat sources detected in the last 7 days!</strong></p>
                        <p>Top threats requiring immediate attention:</p>
                        <ul style="margin: 0;">
                    """.format(critical_threats)
                    
                    for ip in list(critical_ips)[:5]:
                        print """
                            <li><strong>IP {}:</strong> {} attempts targeting {} users (last seen: {})</li>
                        """.format(
                            ip.ClientIp,
                            ip.Attempts,
                            ip.UsersTargeted,
                            format_datetime(ip.LastSeen)
                        )
                    
                    print """
                        </ul>
                    </div>
                    """
                
                # KPI Grid
                print """
                <!-- KPI Grid -->
                <div class="unifsec2025_kpi_grid">
                    <div class="unifsec2025_kpi_card {}">
                        <div class="unifsec2025_kpi_label">Failed Login Attempts</div>
                        <div class="unifsec2025_kpi_value">{:,}</div>
                        <div class="unifsec2025_kpi_subtitle">Last 7 days</div>
                    </div>
                    
                    <div class="unifsec2025_kpi_card {}">
                        <div class="unifsec2025_kpi_label">Unique Threat IPs</div>
                        <div class="unifsec2025_kpi_value">{:,}</div>
                        <div class="unifsec2025_kpi_subtitle">{} critical</div>
                    </div>
                    
                    <div class="unifsec2025_kpi_card {}">
                        <div class="unifsec2025_kpi_label">Users at Risk</div>
                        <div class="unifsec2025_kpi_value">{:,}</div>
                        <div class="unifsec2025_kpi_subtitle">{} high risk</div>
                    </div>
                    
                    <div class="unifsec2025_kpi_card">
                        <div class="unifsec2025_kpi_label">Account Lockouts</div>
                        <div class="unifsec2025_kpi_value">{:,}</div>
                        <div class="unifsec2025_kpi_subtitle">Security triggered</div>
                    </div>
                    
                    <div class="unifsec2025_kpi_card">
                        <div class="unifsec2025_kpi_label">Password Resets</div>
                        <div class="unifsec2025_kpi_value">{:,}</div>
                        <div class="unifsec2025_kpi_subtitle">User initiated</div>
                    </div>
                    
                    <div class="unifsec2025_kpi_card">
                        <div class="unifsec2025_kpi_label">Days with Activity</div>
                        <div class="unifsec2025_kpi_value">{}/7</div>
                        <div class="unifsec2025_kpi_subtitle">{}% coverage</div>
                    </div>
                </div>
                """.format(
                    'critical' if total_attempts > 1000 else 'warning' if total_attempts > 500 else '',
                    total_attempts,
                    'critical' if critical_threats > 0 else 'warning' if unique_threats > 50 else '',
                    unique_threats,
                    critical_threats,
                    'critical' if high_risk_users > 0 else 'warning' if users_affected > 100 else '',
                    users_affected,
                    high_risk_users,
                    lockouts,
                    password_resets,
                    days_active,
                    int((float(days_active) / 7) * 100)
                )
                
            
            # Dashboard Selector
            print """
                <!-- Dashboard Selector -->
                <div class="unifsec2025_dashboard_selector">
                    <h2 style="text-align: center; margin-bottom: 30px;">Select Security Dashboard</h2>
                    
                    <div class="unifsec2025_dashboard_grid">
                        <!-- High-Risk User Monitor -->
                        <div class="unifsec2025_dashboard_card">
                            <div style="font-size: 48px; color: #dc3545; margin-bottom: 15px;">
                                <i class="fa fa-user-shield"></i>
                            </div>
                            <h3 style="color: #212529; margin-bottom: 10px;">High-Risk User Security Monitor</h3>
                            <p style="color: #6c757d; margin-bottom: 20px;">
                                Monitor privileged accounts and high-risk users for potential compromise. 
                                Enhanced tracking for Admin, Finance, Developer, and API accounts.
                            </p>
                            <ul style="text-align: left; color: #495057; font-size: 14px; margin-bottom: 20px;">
                                <li>Privileged account monitoring</li>
                                <li>Role-based risk analysis</li>
                                <li>Login pattern detection</li>
                                <li>Critical security alerts</li>
                                <li>Targeted attack analysis</li>
                            </ul>
                            <form method="post" action="#" class="unifsec2025_dashboard_form">
                                <input type="hidden" name="view" value="highrisk">
                                <button type="submit" class="unifsec2025_dashboard_btn highrisk">
                                    <i class="fa fa-arrow-right"></i> Open High-Risk Monitor
                                </button>
                            </form>
                        </div>
                        
                        <!-- Enhanced Security Dashboard -->
                        <div class="unifsec2025_dashboard_card">
                            <div style="font-size: 48px; color: #007bff; margin-bottom: 15px;">
                                <i class="fa fa-shield"></i>
                            </div>
                            <h3 style="color: #212529; margin-bottom: 10px;">Enhanced Security Dashboard</h3>
                            <p style="color: #6c757d; margin-bottom: 20px;">
                                Advanced threat detection and monitoring with real-time analysis capabilities.
                                Comprehensive security reporting with pattern detection and threat intelligence.
                            </p>
                            <ul style="text-align: left; color: #495057; font-size: 14px; margin-bottom: 20px;">
                                <li>Critical security alerts</li>
                                <li>Advanced threat analysis</li>
                                <li>IP reputation tracking</li>
                                <li>User behavior analysis</li>
                                <li>Compliance reporting</li>
                            </ul>
                            <form method="post" action="#" class="unifsec2025_dashboard_form">
                                <input type="hidden" name="view" value="enhanced">
                                <button type="submit" class="unifsec2025_dashboard_btn enhanced">
                                    <i class="fa fa-arrow-right"></i> Open Enhanced Dashboard
                                </button>
                            </form>
                        </div>
                    </div>
                </div>
                
                <!-- Latest Activity Info -->
                <div style="text-align: center; color: #6c757d; margin-top: 20px;">
                    <p><i class="fa fa-clock-o"></i> Latest security event: {}</p>
                </div>
            </div>
            
            <script>
                // Helper function to get the correct form submission URL
                function getPyScriptAddress() {{
                    let path = window.location.pathname;
                    return path.replace("/PyScript/", "/PyScriptForm/");
                }}
                
                // Set all form actions to correct URL
                document.querySelectorAll('form').forEach(function(form) {{
                    form.action = getPyScriptAddress();
                }});
            </script>
            """.format(
                format_datetime(latest_activity) if latest_activity else "No recent activity"
            )
            
        except Exception as e:
            print_error("KPI Overview Generation", e)

# ::START:: High-Risk User Monitor
class HighRiskUserMonitor:
    """High-risk user security monitoring and threat detection system"""
    
    def __init__(self):
        # ::STEP:: Initialize Configuration
        self.high_risk_roles = ['Admin', 'Finance', 'FinanceAdmin', 'Developer', 'ApiOnly']
        self.critical_threshold = 5   # Failed attempts for critical alert
        self.high_threshold = 3       # Failed attempts for high alert
        self.lookback_days = 30       # Default monitoring period
        self.max_results = 100        # Maximum results to display
        
    def get_configuration_form(self):
        """Generate configuration form for high-risk user monitoring"""
        try:
            print """
            <style>
            .tphighrisk2025_container {{
                max-width: 1200px;
                margin: 0 auto;
                padding: 20px;
                font-family: Arial, sans-serif;
            }}
            .tphighrisk2025_panel {{
                border: 2px solid #dc3545;
                border-radius: 8px;
                margin: 20px 0;
                background-color: #ffffff;
                box-shadow: 0 4px 8px rgba(220, 53, 69, 0.15);
            }}
            .tphighrisk2025_header {{
                background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
                border-bottom: 1px solid #f1aeb5;
                padding: 15px 20px;
                border-radius: 6px 6px 0 0;
                font-weight: bold;
                font-size: 18px;
                color: #721c24;
            }}
            .tphighrisk2025_body {{
                padding: 25px;
            }}
            .tphighrisk2025_formrow {{
                margin-bottom: 20px;
                display: flex;
                gap: 20px;
                flex-wrap: wrap;
            }}
            .tphighrisk2025_formcol {{
                flex: 1;
                min-width: 250px;
            }}
            .tphighrisk2025_formgroup {{
                margin-bottom: 20px;
            }}
            .tphighrisk2025_label {{
                display: block;
                margin-bottom: 8px;
                font-weight: bold;
                color: #495057;
            }}
            .tphighrisk2025_control {{
                width: 100%;
                padding: 10px 12px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-size: 14px;
                background-color: #fff;
                box-sizing: border-box;
            }}
            .tphighrisk2025_control:focus {{
                outline: 0;
                border-color: #dc3545;
                box-shadow: 0 0 0 0.2rem rgba(220, 53, 69, 0.25);
            }}
            .tphighrisk2025_checkbox {{
                display: flex;
                align-items: center;
                margin-bottom: 12px;
            }}
            .tphighrisk2025_checkbox input {{
                margin-right: 8px;
            }}
            .tphighrisk2025_btn {{
                display: inline-block;
                padding: 12px 24px;
                font-size: 16px;
                font-weight: bold;
                text-decoration: none;
                border: none;
                border-radius: 6px;
                cursor: pointer;
                background-color: #dc3545;
                color: #ffffff;
                transition: background-color 0.3s;
                margin: 0 10px;
            }}
            .tphighrisk2025_btn:hover {{
                background-color: #c82333;
            }}
            .tphighrisk2025_btn_secondary {{
                background-color: #6c757d;
            }}
            .tphighrisk2025_btn_secondary:hover {{
                background-color: #5a6268;
            }}
            .tphighrisk2025_alert {{
                padding: 15px;
                margin-bottom: 20px;
                border: 1px solid transparent;
                border-radius: 4px;
                background-color: #fff3cd;
                border-color: #ffeaa7;
                color: #856404;
            }}
            .tphighrisk2025_btncontainer {{
                text-align: center;
                margin-top: 30px;
            }}
            .tphighrisk2025_helptext {{
                font-size: 12px;
                color: #6c757d;
                margin-top: 5px;
            }}
            .tphighrisk2025_loading {{
                display: none;
                margin-left: 15px;
                color: #6c757d;
            }}
            </style>
            
            <div class="tphighrisk2025_container">
                <h2 style="color: #dc3545;"><i class="fa fa-user-shield"></i> High-Risk User Security Monitor</h2>
                <div class="tphighrisk2025_alert">
                    <strong>Privileged Account Protection:</strong> Monitor security events for users with elevated privileges 
                    including Admin, Finance, Developer, and API access roles. Enhanced threat detection for high-value targets.
                </div>
                
                <form method="post" action="#">
                    <input type="hidden" name="action" value="generate_highrisk_report">
                    <input type="hidden" name="view" value="highrisk">
                    
                    <div class="tphighrisk2025_panel">
                        <div class="tphighrisk2025_header">
                            <i class="fa fa-cog"></i> Monitoring Configuration
                        </div>
                        <div class="tphighrisk2025_body">
                            <div class="tphighrisk2025_formrow">
                                <div class="tphighrisk2025_formcol">
                                    <div class="tphighrisk2025_formgroup">
                                        <label class="tphighrisk2025_label" for="lookback_days">Monitoring Period:</label>
                                        <select class="tphighrisk2025_control" name="lookback_days" id="lookback_days">
                                            <option value="1">Last 24 Hours</option>
                                            <option value="3">Last 3 Days</option>
                                            <option value="7">Last 7 Days</option>
                                            <option value="14">Last 2 Weeks</option>
                                            <option value="30" selected>Last 30 Days</option>
                                            <option value="60">Last 2 Months</option>
                                            <option value="90">Last 3 Months</option>
                                        </select>
                                    </div>
                                </div>
                                
                                <div class="tphighrisk2025_formcol">
                                    <div class="tphighrisk2025_formgroup">
                                        <label class="tphighrisk2025_label" for="max_results">Maximum Results:</label>
                                        <select class="tphighrisk2025_control" name="max_results" id="max_results">
                                            <option value="25">Top 25 Users</option>
                                            <option value="50">Top 50 Users</option>
                                            <option value="100" selected>Top 100 Users</option>
                                            <option value="250">Top 250 Users</option>
                                            <option value="0">All High-Risk Users</option>
                                        </select>
                                        <div class="tphighrisk2025_helptext">Limit results for performance</div>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="tphighrisk2025_formrow">
                                <div class="tphighrisk2025_formcol">
                                    <div class="tphighrisk2025_formgroup">
                                        <label class="tphighrisk2025_label" for="critical_threshold">Critical Alert Threshold:</label>
                                        <input type="number" class="tphighrisk2025_control" name="critical_threshold" 
                                               id="critical_threshold" value="5" min="1" max="50">
                                        <div class="tphighrisk2025_helptext">Failed attempts to trigger critical alert</div>
                                    </div>
                                </div>
                                
                                <div class="tphighrisk2025_formcol">
                                    <div class="tphighrisk2025_formgroup">
                                        <label class="tphighrisk2025_label" for="high_threshold">High Alert Threshold:</label>
                                        <input type="number" class="tphighrisk2025_control" name="high_threshold" 
                                               id="high_threshold" value="3" min="1" max="25">
                                        <div class="tphighrisk2025_helptext">Failed attempts to trigger high alert</div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="tphighrisk2025_panel">
                        <div class="tphighrisk2025_header">
                            <i class="fa fa-users"></i> High-Risk Role Configuration
                        </div>
                        <div class="tphighrisk2025_body">
                            <p style="margin-bottom: 20px; color: #495057;">
                                Select which roles should be considered high-risk for enhanced monitoring:
                            </p>
                            <div class="tphighrisk2025_formrow">
                                <div class="tphighrisk2025_formcol">
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="monitor_admin" value="1" checked>
                                        <label><strong>Admin</strong> - System administrators</label>
                                    </div>
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="monitor_finance" value="1" checked>
                                        <label><strong>Finance</strong> - Financial access users</label>
                                    </div>
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="monitor_financeadmin" value="1" checked>
                                        <label><strong>FinanceAdmin</strong> - Finance administrators</label>
                                    </div>
                                </div>
                                
                                <div class="tphighrisk2025_formcol">
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="monitor_developer" value="1" checked>
                                        <label><strong>Developer</strong> - Development access</label>
                                    </div>
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="monitor_apionly" value="1" checked>
                                        <label><strong>ApiOnly</strong> - API access accounts</label>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="tphighrisk2025_formgroup">
                                <label class="tphighrisk2025_label" for="custom_roles">Custom High-Risk Roles (comma-separated):</label>
                                <input type="text" class="tphighrisk2025_control" name="custom_roles" 
                                       id="custom_roles" placeholder="e.g., SuperUser, DBAdmin, SecurityAdmin">
                                <div class="tphighrisk2025_helptext">Additional roles to monitor beyond the standard high-risk roles</div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="tphighrisk2025_panel">
                        <div class="tphighrisk2025_header">
                            <i class="fa fa-list-alt"></i> Report Options
                        </div>
                        <div class="tphighrisk2025_body">
                            <div class="tphighrisk2025_formrow">
                                <div class="tphighrisk2025_formcol">
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="show_alerts" value="1" checked>
                                        <label><strong>Critical Alerts & Notifications</strong></label>
                                    </div>
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="show_user_overview" value="1" checked>
                                        <label><strong>High-Risk User Overview</strong></label>
                                    </div>
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="show_security_events" value="1" checked>
                                        <label><strong>Security Events Analysis</strong></label>
                                    </div>
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="show_role_analysis" value="1" checked>
                                        <label><strong>Role-Based Risk Analysis</strong></label>
                                    </div>
                                </div>
                                
                                <div class="tphighrisk2025_formcol">
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="show_login_patterns" value="1">
                                        <label>Login Pattern Analysis</label>
                                    </div>
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="show_ip_analysis" value="1">
                                        <label>IP Address Analysis</label>
                                    </div>
                                    <div class="tphighrisk2025_checkbox">
                                        <input type="checkbox" name="show_recommendations" value="1" checked>
                                        <label>Security Recommendations</label>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="tphighrisk2025_btncontainer">
                        <button type="submit" class="tphighrisk2025_btn" onclick="showLoading()">
                            <i class="fa fa-shield"></i> Generate High-Risk User Report
                        </button>
                        <span id="tphighrisk2025_loading" class="tphighrisk2025_loading">
                            <i class="fa fa-spinner fa-spin"></i> Analyzing security data...
                        </span>
                    </div>
                </form>
            </div>
            
            <script>
                // Helper function to get the correct form submission URL
                function getPyScriptAddress() {{
                    let path = window.location.pathname;
                    return path.replace("/PyScript/", "/PyScriptForm/");
                }}
                
                // Set form action to correct URL
                document.querySelector('form').action = getPyScriptAddress();
                
                // Show loading indicator on form submit
                document.querySelector('form').addEventListener('submit', function(e) {{
                    document.getElementById('tpsecuritydash2025_loading').style.display = 'inline-block';
                    // Don't disable the button until after form data is collected
                    setTimeout(function() {{
                        document.querySelector('button[type="submit"]').disabled = true;
                    }}, 100);
                }});
                
                // Emergency report function
                function generateEmergencyReport() {{
                    if (confirm('Generate an emergency threat scan for the last 24 hours?')) {{
                        // Set emergency parameters
                        document.getElementById('lookback_days').value = '1';
                        document.getElementById('min_attempts').value = '3';
                        document.getElementById('high_risk_threshold').value = '10';
                        document.getElementById('time_window').value = '15';
                        
                        // Enable critical sections
                        document.querySelector('input[name="show_critical_alerts"]').checked = true;
                        document.querySelector('input[name="show_recent_activity"]').checked = true;
                        document.querySelector('input[name="show_threat_analysis"]').checked = true;
                        
                        // Submit form
                        document.querySelector('form').submit();
                    }}
                }}
            </script>
            """
            
        except Exception as e:
            print_error("Configuration Form Generation", e)
    
    def generate_highrisk_report(self, params):
        """Generate comprehensive high-risk user security report"""
        try:
            # ::STEP:: Process Parameters
            lookback_days = int(params.get('lookback_days', 30))
            max_results = int(params.get('max_results', 100))
            critical_threshold = int(params.get('critical_threshold', 5))
            high_threshold = int(params.get('high_threshold', 3))
            
            # Role monitoring settings
            monitor_admin = params.get('monitor_admin') == '1'
            monitor_finance = params.get('monitor_finance') == '1'
            monitor_financeadmin = params.get('monitor_financeadmin') == '1'
            monitor_developer = params.get('monitor_developer') == '1'
            monitor_apionly = params.get('monitor_apionly') == '1'
            custom_roles = params.get('custom_roles', '')
            
            # Report sections
            show_alerts = params.get('show_alerts') == '1'
            show_user_overview = params.get('show_user_overview') == '1'
            show_security_events = params.get('show_security_events') == '1'
            show_role_analysis = params.get('show_role_analysis') == '1'
            show_login_patterns = params.get('show_login_patterns') == '1'
            show_ip_analysis = params.get('show_ip_analysis') == '1'
            show_recommendations = params.get('show_recommendations') == '1'
            
            # Build monitored roles list
            monitored_roles = []
            if monitor_admin: monitored_roles.append('Admin')
            if monitor_finance: monitored_roles.append('Finance')
            if monitor_financeadmin: monitored_roles.append('FinanceAdmin')
            if monitor_developer: monitored_roles.append('Developer')
            if monitor_apionly: monitored_roles.append('ApiOnly')
            if custom_roles:
                custom_role_list = [role.strip() for role in custom_roles.split(',') if role.strip()]
                monitored_roles.extend(custom_role_list)
            
            # ::STEP:: Generate Report Header
            self.generate_report_header(lookback_days, monitored_roles, critical_threshold, high_threshold)
            
            # ::STEP:: Generate Report Sections
            if show_alerts:
                print '<div class="tphighrisk2025_section">'
                self.generate_highrisk_critical_alerts(lookback_days, critical_threshold, high_threshold, monitored_roles)
                print '</div>'
            
            if show_user_overview:
                print '<div class="tphighrisk2025_section">'
                self.generate_highrisk_user_overview(lookback_days, max_results, monitored_roles)
                print '</div>'
            
            if show_security_events:
                print '<div class="tphighrisk2025_section">'
                self.generate_highrisk_security_events(lookback_days, max_results, critical_threshold, monitored_roles)
                print '</div>'
            
            if show_role_analysis:
                print '<div class="tphighrisk2025_section">'
                # FIXED: Now properly passing lookback_days parameter
                self.generate_role_based_analysis(lookback_days, monitored_roles)
                print '</div>'
            
            if show_login_patterns:
                print '<div class="tphighrisk2025_section">'
                self.generate_login_pattern_analysis(lookback_days, monitored_roles)
                print '</div>'
            
            if show_ip_analysis:
                print '<div class="tphighrisk2025_section">'
                self.generate_highrisk_ip_analysis(lookback_days, monitored_roles)
                print '</div>'
            
            if show_recommendations:
                print '<div class="tphighrisk2025_section">'
                self.generate_highrisk_recommendations(lookback_days, critical_threshold, monitored_roles)
                print '</div>'
            
            # ::STEP:: Generate Report Footer
            self.generate_report_footer()
            
        except Exception as e:
            print_error("High-Risk User Report Generation", e)
    
    def generate_report_header(self, lookback_days, monitored_roles, critical_threshold, high_threshold):
        """Generate report header with styling"""
        print """
        <div class="container-fluid">
            <div class="row">
                <div class="col-md-12">
                    <div style="border-bottom: 3px solid #dc3545; padding-bottom: 20px; margin-bottom: 30px; background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%); padding: 20px; border-radius: 8px;">
                        <h1 style="color: #721c24; font-size: 32px; margin: 0; font-weight: bold;">
                            <i class="fa fa-user-shield"></i> High-Risk User Security Report
                        </h1>
                        <p style="color: #721c24; margin: 15px 0 0 0; font-size: 16px;">
                            <strong>Monitoring Period:</strong> Last {} days | 
                            <strong>Roles Monitored:</strong> {} | 
                            <strong>Alert Thresholds:</strong> Critical: {}, High: {} | 
                            <strong>Generated:</strong> {}
                        </p>
                    </div>
                    
                    <style type="text/css">
                    /* High-Risk User Monitor Styles */
                    .tphighrisk2025_section {{
                        margin-bottom: 40px !important;
                    }}
                    .tphighrisk2025_critical {{
                        background: linear-gradient(135deg, #dc3545 0%, #c82333 100%) !important;
                        color: #ffffff !important;
                        padding: 20px !important;
                        border-radius: 8px !important;
                        margin: 20px 0 !important;
                        box-shadow: 0 4px 8px rgba(220, 53, 69, 0.3) !important;
                        animation: pulse 2s infinite !important;
                    }}
                    @keyframes pulse {{
                        0% {{ box-shadow: 0 0 0 0 rgba(220, 53, 69, 0.7); }}
                        70% {{ box-shadow: 0 0 0 10px rgba(220, 53, 69, 0); }}
                        100% {{ box-shadow: 0 0 0 0 rgba(220, 53, 69, 0); }}
                    }}
                    .tphighrisk2025_high {{
                        background: linear-gradient(135deg, #ffc107 0%, #e0a800 100%) !important;
                        color: #000000 !important;
                        padding: 15px !important;
                        border-radius: 6px !important;
                        margin: 15px 0 !important;
                    }}
                    .tphighrisk2025_clickable {{
                        cursor: pointer !important;
                        color: #007bff !important;
                        text-decoration: underline !important;
                        border: 1px solid transparent !important;
                        padding: 3px 6px !important;
                        border-radius: 4px !important;
                        transition: all 0.2s ease !important;
                    }}
                    .tphighrisk2025_clickable:hover {{
                        background-color: #e3f2fd !important;
                        border-color: #007bff !important;
                        color: #0056b3 !important;
                    }}
                    .tphighrisk2025_details {{
                        display: none;
                        margin-top: 15px;
                        padding: 15px;
                        background-color: #f8f9fa;
                        border: 1px solid #dee2e6;
                        border-radius: 6px;
                        border-left: 4px solid #dc3545;
                    }}
                    .tphighrisk2025_rolehigh {{
                        background-color: #dc3545 !important;
                        color: #ffffff !important;
                        padding: 3px 8px !important;
                        border-radius: 4px !important;
                        font-size: 11px !important;
                        font-weight: bold !important;
                    }}
                    .tphighrisk2025_rolemedium {{
                        background-color: #ffc107 !important;
                        color: #000000 !important;
                        padding: 3px 8px !important;
                        border-radius: 4px !important;
                        font-size: 11px !important;
                        font-weight: bold !important;
                    }}
                    .tphighrisk2025_rolelow {{
                        background-color: #28a745 !important;
                        color: #ffffff !important;
                        padding: 3px 8px !important;
                        border-radius: 4px !important;
                        font-size: 11px !important;
                        font-weight: bold !important;
                    }}
                    </style>
                    
                    <script>
                    function toggleDetails(id) {{
                        var element = document.getElementById(id);
                        if (element.style.display === "none" || element.style.display === "") {{
                            element.style.display = "block";
                        }} else {{
                            element.style.display = "none";
                        }}
                    }}
                    
                    function viewUser(userId) {{
                        if (userId && userId !== 'NULL' && userId !== '0') {{
                            window.open('/Person2/' + userId, '_blank');
                        }} else {{
                            alert('User ID not available');
                        }}
                    }}
                    
                    function openIPLookup(ip) {{
                        if (ip) {{
                            var url = 'https://www.abuseipdb.com/check/' + ip;
                            window.open(url, '_blank');
                        }}
                    }}
                    </script>
        """.format(
            lookback_days,
            ', '.join(monitored_roles) if monitored_roles else 'All Standard Roles',
            critical_threshold,
            high_threshold,
            datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )
    
    def generate_highrisk_critical_alerts(self, lookback_days, critical_threshold, high_threshold, monitored_roles):
        """Generate critical alerts for high-risk users"""
        try:
            # Build role condition for SQL
            if monitored_roles:
                role_conditions = []
                for role in monitored_roles:
                    role_conditions.append("ul.Roles LIKE '%{}%'".format(role))
                role_where = " AND ({})".format(" OR ".join(role_conditions))
            else:
                role_where = " AND (ul.Roles LIKE '%Admin%' OR ul.Roles LIKE '%Finance%' OR ul.Roles LIKE '%Developer%' OR ul.Roles LIKE '%ApiOnly%')"
            
            # ::STEP:: Query Critical Alerts with Users table integration
            sql_critical = """
                SELECT 
                    ul.Username,
                    ul.PeopleId,
                    ul.Roles,
                    u.IsLockedOut,
                    u.LastLockedOutDate,
                    u.FailedPasswordAttemptCount AS CurrentFailedAttempts,
                    u.MustChangePassword,
                    u.LastLoginDate,
                    COUNT(*) AS FailedAttempts,
                    COUNT(DISTINCT al.ClientIp) AS UniqueIPs,
                    COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) AS DaysWithFailures,
                    MIN(al.ActivityDate) AS FirstFailure,
                    MAX(al.ActivityDate) AS LastFailure,
                    SUM(CASE WHEN al.Activity LIKE '%locked%' THEN 1 ELSE 0 END) AS ActivityLogLockouts,
                    SUM(CASE WHEN al.Activity LIKE '%passwordreset%' OR al.Activity LIKE '%ForgotPassword%' THEN 1 ELSE 0 END) AS PasswordResetAttempts
                FROM ActivityLog al
                JOIN UserList ul ON al.UserId = ul.UserId
                LEFT JOIN Users u ON al.UserId = u.UserId
                WHERE al.Activity LIKE '%failed%'
                    AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                    {}
                GROUP BY ul.Username, ul.PeopleId, ul.Roles, u.IsLockedOut, 
                         u.LastLockedOutDate, u.FailedPasswordAttemptCount, 
                         u.MustChangePassword, u.LastLoginDate
                HAVING COUNT(*) >= {}
                ORDER BY 
                    CASE WHEN u.IsLockedOut = 1 THEN 0 ELSE 1 END,  -- Locked accounts first
                    FailedAttempts DESC;
            """.format(lookback_days, role_where, high_threshold)
            
            critical_data = q.QuerySql(sql_critical)
            
            if critical_data and len(critical_data) > 0:
                # Separate critical and high alerts
                critical_users = []
                high_users = []
                
                for user in critical_data:
                    if user.FailedAttempts >= critical_threshold:
                        critical_users.append(user)
                    else:
                        high_users.append(user)
                
                if critical_users:
                    print """
                    <div class="tphighrisk2025_critical">
                        <h2 style="margin-top: 0; font-size: 24px;">
                            <i class="fa fa-exclamation-triangle fa-2x"></i> CRITICAL ALERT - HIGH-RISK USER COMPROMISE
                        </h2>
                        <p style="font-size: 18px; font-weight: bold;">IMMEDIATE ACTION REQUIRED - PRIVILEGED ACCOUNTS UNDER ATTACK</p>
                        <div style="background-color: rgba(255,255,255,0.2); padding: 15px; border-radius: 6px; margin: 15px 0;">
                            <h4 style="margin-top: 0;">Compromised High-Risk Accounts:</h4>
                            <ul style="font-size: 16px; margin: 0;">
                    """
                    
                    for user in critical_users[:5]:
                        roles_display = format_roles_display(user.Roles, short=True)
                        print """
                                <li><strong>{}</strong> [{}]: {} failed attempts from {} IPs</li>
                        """.format(user.Username, roles_display, user.FailedAttempts, user.UniqueIPs)
                    
                    print """
                            </ul>
                            <h4 style="margin: 20px 0 10px 0;">EMERGENCY RESPONSE REQUIRED:</h4>
                            <ol style="font-size: 14px; margin: 0;">
                                <li><strong>DISABLE</strong> affected high-privilege accounts immediately</li>
                                <li><strong>FORCE</strong> password reset for all compromised accounts</li>
                                <li><strong>AUDIT</strong> recent activities and system changes</li>
                                <li><strong>BLOCK</strong> attacking IP addresses at firewall level</li>
                                <li><strong>ENABLE</strong> multi-factor authentication if not already active</li>
                                <li><strong>NOTIFY</strong> security team and executive management</li>
                                <li><strong>INITIATE</strong> incident response procedures</li>
                            </ol>
                        </div>
                    </div>
                    """
                
                if high_users:
                    print """
                    <div class="tphighrisk2025_high">
                        <h3 style="margin-top: 0;">
                            <i class="fa fa-exclamation-circle"></i> HIGH PRIORITY ALERT - Privileged Account Security Events
                        </h3>
                        <p><strong>High-risk users with security incidents requiring immediate attention:</strong></p>
                        <ul style="margin: 10px 0;">
                    """
                    
                    for user in high_users[:10]:
                        roles_display = format_roles_display(user.Roles, short=True)
                        print """
                            <li><strong>{}</strong> [{}]: {} failed attempts</li>
                        """.format(user.Username, roles_display, user.FailedAttempts)
                    
                    print """
                        </ul>
                        <p><strong>Actions Required:</strong> Review accounts, verify legitimate access, consider temporary restrictions.</p>
                    </div>
                    """
            else:
                print """
                <div style="border: 2px solid #28a745; background-color: #d4edda; color: #155724; padding: 20px; margin: 20px 0; border-radius: 8px;">
                    <h3 style="color: #155724; margin-top: 0;">
                        <i class="fa fa-check-circle"></i> No Critical High-Risk User Threats
                    </h3>
                    <p style="margin: 0;">No high-risk users have security incidents exceeding alert thresholds in the monitoring period.</p>
                </div>
                """
                
        except Exception as e:
            print_error("High-Risk Critical Alerts Generation", e)
    
    def generate_highrisk_user_overview(self, lookback_days, max_results, monitored_roles):
        """Generate overview of all high-risk users and their security status"""
        try:
            # Build role condition
            if monitored_roles:
                role_conditions = []
                for role in monitored_roles:
                    role_conditions.append("ul.Roles LIKE '%{}%'".format(role))
                role_where = " AND ({})".format(" OR ".join(role_conditions))
            else:
                role_where = " AND (ul.Roles LIKE '%Admin%' OR ul.Roles LIKE '%Finance%' OR ul.Roles LIKE '%Developer%' OR ul.Roles LIKE '%ApiOnly%')"
            
            # ::STEP:: Query User Overview
            limit_clause = "TOP {}".format(max_results) if max_results > 0 else ""
            
            sql_overview = """
                SELECT {}
                    ul.Username,
                    ul.PeopleId,
                    ul.Roles,
                    ISNULL(security_events.FailedAttempts, 0) AS FailedAttempts,
                    ISNULL(security_events.UniqueIPs, 0) AS UniqueIPs,
                    ISNULL(security_events.DaysWithFailures, 0) AS DaysWithFailures,
                    security_events.LastFailure,
                    security_events.AccountLockouts,
                    CASE 
                        WHEN ISNULL(security_events.FailedAttempts, 0) >= 5 THEN 'Critical Risk'
                        WHEN ISNULL(security_events.FailedAttempts, 0) >= 3 THEN 'High Risk'
                        WHEN ISNULL(security_events.FailedAttempts, 0) >= 1 THEN 'Medium Risk'
                        ELSE 'Low Risk'
                    END AS RiskLevel
                FROM UserList ul
                LEFT JOIN (
                    SELECT 
                        al.UserId,
                        COUNT(*) AS FailedAttempts,
                        COUNT(DISTINCT al.ClientIp) AS UniqueIPs,
                        COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) AS DaysWithFailures,
                        MAX(al.ActivityDate) AS LastFailure,
                        SUM(CASE WHEN al.Activity LIKE '%locked%' THEN 1 ELSE 0 END) AS AccountLockouts
                    FROM ActivityLog al
                    WHERE al.Activity LIKE '%failed%'
                        AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                    GROUP BY al.UserId
                ) security_events ON ul.UserId = security_events.UserId
                WHERE ul.Username IS NOT NULL
                    {}
                ORDER BY ISNULL(security_events.FailedAttempts, 0) DESC, ul.Username;
            """.format(limit_clause, lookback_days, role_where)
            
            overview_data = q.QuerySql(sql_overview)
            
            if overview_data and len(overview_data) > 0:
                print """
<div style="border: 2px solid #dc3545; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
    <div style="background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%); border-bottom: 1px solid #f1aeb5; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #721c24;">
        <i class="fa fa-users"></i> High-Risk User Security Overview
        <span style="float: right; font-size: 14px; font-weight: normal;">
            Monitoring {} privileged accounts
        </span>
    </div>
    <div style="padding: 20px; overflow-x: auto;">
        <table style="width: 100%; border-collapse: collapse; margin: 0; font-size: 13px;">
            <thead>
                <tr>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Username</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Roles</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Risk Level</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Failed Attempts</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Source IPs</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Days Active</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Last Event</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Actions</th>
                </tr>
            </thead>
            <tbody>
                """.format(len(list(overview_data)) if hasattr(overview_data, '__len__') else "multiple")
                
                row_counter = 0
                for user in overview_data:
                    row_counter += 1
                    
                    risk_level = getattr(user, 'RiskLevel', 'Low Risk')
                    failed_attempts = getattr(user, 'FailedAttempts', 0)
                    unique_ips = getattr(user, 'UniqueIPs', 0)
                    days_with_failures = getattr(user, 'DaysWithFailures', 0)
                    last_failure = getattr(user, 'LastFailure', None)
                    account_lockouts = getattr(user, 'AccountLockouts', 0)
                    
                    # Risk level styling
                    if 'Critical' in risk_level:
                        row_bg = 'background-color: #f8d7da;'
                        risk_style = 'tphighrisk2025_rolehigh'
                    elif 'High' in risk_level:
                        row_bg = 'background-color: #fff3cd;'
                        risk_style = 'tphighrisk2025_rolemedium'
                    elif 'Medium' in risk_level:
                        row_bg = ''
                        risk_style = 'tphighrisk2025_rolemedium'
                    else:
                        row_bg = ''
                        risk_style = 'tphighrisk2025_rolelow'
                    
                    roles_display = format_roles_display(user.Roles)
                    
                    print """
                    <tr style="{}">
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <a href="javascript:viewUser({})" style="color: #dc3545; font-weight: bold; text-decoration: underline;">
                                {} <i class="fa fa-external-link"></i>
                            </a>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            {}
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="{}">{}</span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">
                            <span class="tphighrisk2025_clickable" onclick="toggleDetails('user_attempts_{}')">
                                <strong>{}</strong> <i class="fa fa-caret-down"></i>
                            </span>
                            <div id="user_attempts_{}" class="tphighrisk2025_details">
                                <strong>Security Event Details for {}:</strong><br>
                                • <strong>Failed Attempts:</strong> {}<br>
                                • <strong>Account Lockouts:</strong> {}<br>
                                • <strong>Attack Sources:</strong> {} unique IPs<br>
                                • <strong>Time Span:</strong> {} days with incidents<br>
                                • <strong>Last Event:</strong> {}
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; font-size: 11px;">
                            {}
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">
                            <button onclick="viewUser({})" style="padding: 3px 8px; font-size: 10px; background-color: #dc3545; color: white; border: none; border-radius: 3px; cursor: pointer; margin: 1px;">
                                <i class="fa fa-user"></i>
                            </button>
                        </td>
                    </tr>
                    """.format(
                        row_bg,
                        user.PeopleId, user.Username,
                        roles_display,
                        risk_style, risk_level,
                        row_counter,
                        failed_attempts,
                        row_counter,
                        user.Username,
                        failed_attempts,
                        account_lockouts,
                        unique_ips,
                        days_with_failures,
                        format_datetime(last_failure) if last_failure else 'No recent events',
                        unique_ips,
                        days_with_failures,
                        format_datetime_short(last_failure) if last_failure else 'No recent events',
                        user.PeopleId
                    )
                
                print """
            </tbody>
        </table>
        
        <div style="margin-top: 20px; padding: 18px; background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-radius: 6px; border-left: 4px solid #ffc107;">
            <strong><i class="fa fa-info-circle"></i> High-Risk User Monitoring:</strong> 
            <ul style="margin-top: 10px; margin-bottom: 0; font-size: 13px;">
                <li><strong>Privileged Account Protection:</strong> Enhanced monitoring for users with elevated system access</li>
                <li><strong>Risk Classification:</strong> Automatic risk scoring based on failed authentication attempts</li>
                <li><strong>Real-time Alerts:</strong> Immediate notifications for critical security events</li>
                <li><strong>Interactive Analysis:</strong> Click usernames to view profiles, click attempts for detailed breakdowns</li>
                <li><strong>Continuous Monitoring:</strong> 24/7 surveillance of high-value target accounts</li>
            </ul>
        </div>
    </div>
</div>
                """
            else:
                print """
<div style="border: 2px solid #28a745; background-color: #d4edda; color: #155724; padding: 15px; margin: 20px 0; border-radius: 6px;">
    <strong>No High-Risk Users Found:</strong> No users with the specified high-risk roles were found in the system.
</div>
                """
                
        except Exception as e:
            print_error("High-Risk User Overview Generation", e)

    def display_role_analysis_results(self, role_name, role_stats, lookback_days):
        """Display results for a single role analysis"""
        try:
            total_users_with_role = getattr(role_stats, 'TotalUsersWithRole', 0)
            users_with_failures = getattr(role_stats, 'UsersWithFailures', 0)
            total_failures = getattr(role_stats, 'TotalFailures', 0)
            unique_ips = getattr(role_stats, 'UniqueIPs', 0)
            days_with_incidents = getattr(role_stats, 'DaysWithIncidents', 0)
            currently_locked = getattr(role_stats, 'CurrentlyLockedUsers', 0)
            recently_locked = getattr(role_stats, 'RecentlyLockedUsers', 0)
            forced_resets = getattr(role_stats, 'ForcedPasswordResets', 0)
            
            # Calculate metrics
            if users_with_failures > 0:
                avg_failed_calc = float(total_failures) / float(users_with_failures)
                avg_failed_str = "{:.1f}".format(avg_failed_calc)
            else:
                avg_failed_str = "0.0"
            
            # Calculate percentage of role users with security incidents
            if total_users_with_role > 0:
                incident_rate = float(users_with_failures) / float(total_users_with_role) * 100
                incident_rate_str = "{:.1f}%".format(incident_rate)
            else:
                incident_rate_str = "0.0%"
            
            # Determine risk color based on multiple factors
            if currently_locked > 0 or total_failures >= 20:
                risk_color = '#dc3545'
                risk_level = 'Critical Risk'
            elif recently_locked > 0 or total_failures >= 10:
                risk_color = '#ffc107'
                risk_level = 'High Risk'
            elif total_failures >= 5:
                risk_color = '#fd7e14'
                risk_level = 'Medium Risk'
            else:
                risk_color = '#28a745'
                risk_level = 'Low Risk'
            
            print """
            <div style="margin: 15px 0; padding: 15px; border: 1px solid {}; border-radius: 6px; background-color: #f8f9fa;">
                <h4 style="color: {}; margin-top: 0;">
                    {} Role Security Analysis
                    <span style="float: right; background-color: {}; color: white; padding: 4px 8px; border-radius: 4px; font-size: 12px;">
                        {}
                    </span>
                </h4>
                <div class="row">
                    <div class="col-md-3">
                        <strong>Total {} Users:</strong> {}
                    </div>
                    <div class="col-md-3">
                        <strong>Users with Incidents:</strong> {} ({})
                    </div>
                    <div class="col-md-3">
                        <strong>Security Events:</strong> {}
                    </div>
                    <div class="col-md-3">
                        <strong>Attack Sources:</strong> {}
                    </div>
                </div>
                <div class="row" style="margin-top: 10px;">
                    <div class="col-md-3">
                        <strong>Currently Locked:</strong> <span style="color: {};">{}</span>
                    </div>
                    <div class="col-md-3">
                        <strong>Recently Locked:</strong> <span style="color: {};">{}</span>
                    </div>
                    <div class="col-md-3">
                        <strong>Forced Resets:</strong> <span style="color: {};">{}</span>
                    </div>
                    <div class="col-md-3">
                        <strong>Avg Failed/Affected User:</strong> {}
                    </div>
                </div>
                <div style="margin-top: 10px; padding: 8px; background-color: rgba(0,0,0,0.05); border-radius: 4px; font-size: 13px;">
                    <strong>Period Analysis:</strong> Over the last {} days, {} out of {} {} users ({}) experienced security incidents, 
                    generating {} total failed authentication attempts from {} unique IP sources across {} days.
                </div>
            </div>
            """.format(
                risk_color, risk_color, role_name,
                risk_color, risk_level,
                role_name, total_users_with_role,
                users_with_failures, incident_rate_str,
                total_failures, unique_ips,
                '#dc3545' if currently_locked > 0 else '#28a745', currently_locked,
                '#ffc107' if recently_locked > 0 else '#28a745', recently_locked,
                '#fd7e14' if forced_resets > 0 else '#28a745', forced_resets,
                avg_failed_str,
                lookback_days, users_with_failures, total_users_with_role, role_name, incident_rate_str,
                total_failures, unique_ips, days_with_incidents
            )
            
        except Exception as e:
            print """
            <div style="margin: 15px 0; padding: 15px; border: 1px solid #dc3545; border-radius: 6px; background-color: #f8d7da;">
                <h4 style="color: #721c24; margin-top: 0;">
                    {} Role Analysis - Error
                </h4>
                <p style="color: #721c24; margin: 0;">
                    Error displaying role analysis: {}
                </p>
            </div>
            """.format(role_name, str(e)[:100])
            

    def generate_highrisk_security_events(self, lookback_days, max_results, critical_threshold, monitored_roles):
        """Generate detailed security events analysis for high-risk users"""
        try:
            # Build role condition
            if monitored_roles:
                role_conditions = []
                for role in monitored_roles:
                    role_conditions.append("ul.Roles LIKE '%{}%'".format(role))
                role_where = " AND ({})".format(" OR ".join(role_conditions))
            else:
                role_where = " AND (ul.Roles LIKE '%Admin%' OR ul.Roles LIKE '%Finance%' OR ul.Roles LIKE '%Developer%' OR ul.Roles LIKE '%ApiOnly%')"
            
            # ::STEP:: Query Security Events
            limit_clause = "TOP {}".format(max_results) if max_results > 0 else ""
            
            sql_events = """
                SELECT {}
                    al.ActivityDate,
                    al.Activity,
                    al.ClientIp,
                    ul.Username,
                    ul.PeopleId,
                    ul.Roles,
                    al.Machine,
                    CASE 
                        WHEN al.Activity LIKE '%ForgotPassword%' THEN 'Password Reset'
                        WHEN al.Activity LIKE '%invalid log in%' THEN 'Invalid Login'
                        WHEN al.Activity LIKE '%failed password%' THEN 'Failed Password'
                        WHEN al.Activity LIKE '%locked%' THEN 'Account Lockout'
                        ELSE 'Other Auth Failure'
                    END AS EventType,
                    CASE 
                        WHEN ul.Roles LIKE '%Admin%' THEN 'Critical'
                        WHEN ul.Roles LIKE '%Finance%' THEN 'High'
                        WHEN ul.Roles LIKE '%Developer%' THEN 'High'
                        ELSE 'Medium'
                    END AS SeverityLevel
                FROM ActivityLog al
                JOIN UserList ul ON al.UserId = ul.UserId
                WHERE al.Activity LIKE '%failed%'
                    AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                    {}
                ORDER BY al.ActivityDate DESC;
            """.format(limit_clause, lookback_days, role_where)
            
            events_data = q.QuerySql(sql_events)
            
            if events_data and len(events_data) > 0:
                print """
<div style="border: 2px solid #e83e8c; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
    <div style="background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%); border-bottom: 1px solid #f1aeb5; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #721c24;">
        <i class="fa fa-exclamation-triangle"></i> High-Risk User Security Events Timeline
    </div>
    <div style="padding: 20px; max-height: 600px; overflow-y: auto;">
                """
                
                event_count = 0
                current_date = None
                
                for event in events_data:
                    event_count += 1
                    event_date = format_datetime(event.ActivityDate)[:10]
                    event_time = format_time(event.ActivityDate)
                    
                    # Add date separator
                    if current_date != event_date:
                        if current_date is not None:
                            print "</div>"
                        
                        current_date = event_date
                        print """
                        <div style="margin: 20px 0 15px 0; padding: 8px 12px; background-color: #e9ecef; border-left: 4px solid #dc3545; font-weight: bold; color: #495057;">
                            <i class="fa fa-calendar"></i> {}
                        </div>
                        <div style="margin-left: 20px;">
                        """.format(event_date)
                    
                    # Determine severity styling
                    severity = getattr(event, 'SeverityLevel', 'Medium')
                    event_type = getattr(event, 'EventType', 'Unknown')
                    
                    if severity == 'Critical':
                        severity_color = 'background-color: #dc3545; color: #ffffff;'
                        timeline_color = '#dc3545'
                    elif severity == 'High':
                        severity_color = 'background-color: #ffc107; color: #000000;'
                        timeline_color = '#ffc107'
                    else:
                        severity_color = 'background-color: #6c757d; color: #ffffff;'
                        timeline_color = '#6c757d'
                    
                    roles_display = format_roles_display(event.Roles, short=True)
                    
                    print """
                    <div style="margin: 8px 0; padding: 12px; border: 1px solid #dee2e6; border-left: 4px solid {}; border-radius: 4px; background-color: #f8f9fa;">
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <span style="font-weight: bold; color: #721c24; margin-right: 10px;">{}</span>
                            <span style="display: inline-block; padding: 2px 6px; font-size: 11px; font-weight: bold; border-radius: 3px; {}">{}</span>
                            <span style="margin-left: auto; font-size: 12px; color: #6c757d;">Event #{}</span>
                        </div>
                        <div style="font-size: 13px; color: #495057; margin-bottom: 6px;">
                            <strong>User:</strong> 
                            <a href="javascript:viewUser({})" style="color: #dc3545; font-weight: bold; text-decoration: underline;">
                                {}
                            </a>
                            <span style="color: #6c757d;">[{}]</span>
                        </div>
                        <div style="font-size: 13px; color: #495057; margin-bottom: 6px;">
                            <strong>Event:</strong> {}
                        </div>
                        <div style="font-size: 12px; color: #6c757d;">
                            <strong>IP:</strong> 
                            <span class="tphighrisk2025_clickable" onclick="openIPLookup('{}')">
                                {}
                            </span>
                             | <strong>Machine:</strong> {} |
                            <strong>Details:</strong> 
                            <span class="tphighrisk2025_clickable" onclick="toggleDetails('event_{}')">
                                View Details <i class="fa fa-caret-down"></i>
                            </span>
                        </div>
                        <div id="event_{}" class="tphighrisk2025_details">
                            <strong>Complete Security Event Details:</strong><br>
                            <code style="background-color: #e9ecef; padding: 8px; border-radius: 4px; display: block; margin: 8px 0; word-wrap: break-word;">
                                {}
                            </code>
                            <strong>High-Risk User Analysis:</strong><br>
                            • <strong>Username:</strong> {}<br>
                            • <strong>Roles:</strong> {}<br>
                            • <strong>Event Type:</strong> {}<br>
                            • <strong>Severity:</strong> {}<br>
                            • <strong>Source IP:</strong> {}<br>
                            • <strong>Machine:</strong> {}<br>
                            • <strong>Timestamp:</strong> {}
                        </div>
                    </div>
                    """.format(
                        timeline_color,
                        event_time,
                        severity_color, severity,
                        event_count,
                        event.PeopleId, event.Username,
                        roles_display,
                        event_type,
                        event.ClientIp, event.ClientIp,
                        getattr(event, 'Machine', 'Unknown'),
                        event_count,
                        event_count,
                        getattr(event, 'Activity', 'Unknown activity'),
                        event.Username,
                        roles_display,
                        event_type,
                        severity,
                        event.ClientIp,
                        getattr(event, 'Machine', 'Unknown'),
                        format_datetime(event.ActivityDate)
                    )
                
                if current_date is not None:
                    print "</div>"
                
                print """
        <div style="margin-top: 20px; padding: 15px; background-color: #fff3cd; border-radius: 4px; border-left: 4px solid #ffc107;">
            <strong><i class="fa fa-info-circle"></i> High-Risk Security Events:</strong> 
            <ul style="margin-top: 10px; margin-bottom: 0;">
                <li>Events shown in chronological order (most recent first)</li>
                <li>Enhanced monitoring for privileged accounts with elevated system access</li>
                <li>Click usernames to view complete user profiles and access history</li>
                <li>Click IP addresses to perform reputation lookups and threat analysis</li>
                <li>Click "View Details" to see complete activity log entries and analysis</li>
            </ul>
        </div>
    </div>
</div>
                """
            else:
                print """
<div style="border: 2px solid #28a745; background-color: #d4edda; color: #155724; padding: 15px; margin: 20px 0; border-radius: 6px;">
    <strong>No Security Events:</strong> No security incidents found for high-risk users in the monitoring period.
</div>
                """
                
        except Exception as e:
            print_error("High-Risk Security Events Generation", e)
    
    def generate_role_based_analysis(self, lookback_days, monitored_roles):
        """Generate analysis by role type"""
        try:
            print """
    <div style="border: 2px solid #6f42c1; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
        <div style="background: linear-gradient(135deg, #e2d9f3 0%, #d6c7f0 100%); border-bottom: 1px solid #c7b3ed; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #333333;">
            <i class="fa fa-users-cog"></i> Role-Based Security Risk Analysis (Last {} Days)
        </div>
        <div style="padding: 20px;">
            """.format(lookback_days)
            
            # ::STEP:: OPTIMIZED - Single query for all roles to avoid multiple table scans
            if monitored_roles:
                # Build role conditions for a single query
                role_conditions = []
                for role in monitored_roles:
                    role_conditions.append("ul.Roles LIKE '%{}%'".format(role))
                role_where = " OR ".join(role_conditions)
                
                sql_all_roles = """
                    SELECT 
                        -- Extract role name for grouping
                        CASE 
                            WHEN ul.Roles LIKE '%Admin%' THEN 'Admin'
                            WHEN ul.Roles LIKE '%Finance%' THEN 'Finance' 
                            WHEN ul.Roles LIKE '%FinanceAdmin%' THEN 'FinanceAdmin'
                            WHEN ul.Roles LIKE '%Developer%' THEN 'Developer'
                            WHEN ul.Roles LIKE '%ApiOnly%' THEN 'ApiOnly'
                            ELSE 'Other'
                        END AS RoleName,
                        
                        -- Count unique users with this role (total)
                        COUNT(DISTINCT ul.UserId) AS TotalUsersWithRole,
                        
                        -- Count users with failures in the period
                        COUNT(DISTINCT CASE WHEN al.ActivityDate IS NOT NULL THEN ul.UserId END) AS UsersWithFailures,
                        
                        -- Count total failures in the period
                        COUNT(al.ActivityDate) AS TotalFailures,
                        
                        -- Count unique IPs attacking users with this role
                        COUNT(DISTINCT al.ClientIp) AS UniqueIPs,
                        
                        -- Count unique days with incidents
                        COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) AS DaysWithIncidents,
                        
                        -- Current account status (not time-limited)
                        COUNT(DISTINCT CASE WHEN u.IsLockedOut = 1 THEN ul.UserId END) AS CurrentlyLockedUsers,
                        COUNT(DISTINCT CASE WHEN u.LastLockedOutDate >= DATEADD(DAY, -{}, GETDATE()) THEN ul.UserId END) AS RecentlyLockedUsers,
                        COUNT(DISTINCT CASE WHEN u.MustChangePassword = 1 THEN ul.UserId END) AS ForcedPasswordResets
                        
                    FROM UserList ul
                    LEFT JOIN ActivityLog al ON ul.UserId = al.UserId 
                        AND al.Activity LIKE '%failed%'
                        AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                    LEFT JOIN Users u ON ul.UserId = u.UserId
                    WHERE ({})
                    GROUP BY 
                        CASE 
                            WHEN ul.Roles LIKE '%Admin%' THEN 'Admin'
                            WHEN ul.Roles LIKE '%Finance%' THEN 'Finance'
                            WHEN ul.Roles LIKE '%FinanceAdmin%' THEN 'FinanceAdmin' 
                            WHEN ul.Roles LIKE '%Developer%' THEN 'Developer'
                            WHEN ul.Roles LIKE '%ApiOnly%' THEN 'ApiOnly'
                            ELSE 'Other'
                        END
                    ORDER BY TotalFailures DESC;
                """.format(lookback_days, lookback_days, role_where)
                
                try:
                    role_data = q.QuerySql(sql_all_roles)
                    if role_data and len(role_data) > 0:
                        for role_stats in role_data:
                            role_name = getattr(role_stats, 'RoleName', 'Unknown')
                            
                            # Skip if this role wasn't in our monitored list
                            if role_name not in monitored_roles and role_name != 'Other':
                                continue
                                
                            self.display_role_analysis_results(role_name, role_stats, lookback_days)
                    else:
                        print """
                        <div style="margin: 15px 0; padding: 15px; border: 1px solid #28a745; border-radius: 6px; background-color: #d4edda; color: #155724;">
                            <strong>No Security Incidents:</strong> No users with monitored roles experienced security incidents in the last {} days.
                        </div>
                        """.format(lookback_days)
                except Exception as role_error:
                    print """
                    <div style="margin: 15px 0; padding: 15px; border: 1px solid #dc3545; border-radius: 6px; background-color: #f8d7da;">
                        <h4 style="color: #721c24; margin-top: 0;">
                            Role Analysis Error
                        </h4>
                        <p style="color: #721c24; margin: 0;">
                            Error analyzing roles: {}
                        </p>
                    </div>
                    """.format(str(role_error)[:100])
            else:
                print """
                <div style="margin: 15px 0; padding: 15px; border: 1px solid #ffc107; border-radius: 6px; background-color: #fff3cd;">
                    <strong>No Roles Selected:</strong> No high-risk roles selected for monitoring.
                </div>
                """
            
        except Exception as e:
            print_error("Role-Based Analysis Generation", e)
    
    def generate_login_pattern_analysis(self, lookback_days, monitored_roles):
        """Generate login pattern analysis for high-risk users"""
        try:
            # Build role condition
            if monitored_roles:
                role_conditions = []
                for role in monitored_roles:
                    role_conditions.append("ul.Roles LIKE '%{}%'".format(role))
                role_where = " AND ({})".format(" OR ".join(role_conditions))
            else:
                role_where = " AND (ul.Roles LIKE '%Admin%' OR ul.Roles LIKE '%Finance%' OR ul.Roles LIKE '%Developer%' OR ul.Roles LIKE '%ApiOnly%')"
            
            # ::STEP:: Query Login Patterns
            sql_patterns = """
                SELECT 
                    ul.Username,
                    ul.PeopleId,
                    ul.Roles,
                    COUNT(*) AS TotalEvents,
                    COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) AS DaysActive,
                    COUNT(DISTINCT DATEPART(HOUR, al.ActivityDate)) AS UniqueHours,
                    COUNT(DISTINCT al.ClientIp) AS UniqueIPs,
                    MIN(al.ActivityDate) AS FirstEvent,
                    MAX(al.ActivityDate) AS LastEvent,
                    CASE 
                        WHEN COUNT(DISTINCT al.ClientIp) >= 5 THEN 'Multiple IP Pattern'
                        WHEN COUNT(DISTINCT DATEPART(HOUR, al.ActivityDate)) >= 12 THEN 'Extended Time Pattern'
                        WHEN COUNT(*) >= 10 THEN 'High Frequency Pattern'
                        ELSE 'Standard Pattern'
                    END AS PatternType
                FROM ActivityLog al
                JOIN UserList ul ON al.UserId = ul.UserId
                WHERE al.Activity LIKE '%failed%'
                    AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                    {}
                GROUP BY ul.Username, ul.PeopleId, ul.Roles
                HAVING COUNT(*) >= 1
                ORDER BY TotalEvents DESC;
            """.format(lookback_days, role_where)
            
            pattern_data = q.QuerySql(sql_patterns)
            
            if pattern_data and len(pattern_data) > 0:
                print """
<div style="border: 2px solid #17a2b8; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
    <div style="background: linear-gradient(135deg, #d1ecf1 0%, #bee5eb 100%); border-bottom: 1px solid #abdde5; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #333333;">
        <i class="fa fa-chart-line"></i> High-Risk User Login Pattern Analysis
    </div>
    <div style="padding: 20px; overflow-x: auto;">
        <table style="width: 100%; border-collapse: collapse; margin: 0; font-size: 13px;">
            <thead>
                <tr>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">User</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Roles</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Pattern Type</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Events</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Days Active</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Time Diversity</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">IP Diversity</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Time Span</th>
                </tr>
            </thead>
            <tbody>
                """
                
                for pattern in pattern_data:
                    pattern_type = getattr(pattern, 'PatternType', 'Standard Pattern')
                    
                    # Pattern type styling
                    if 'Multiple IP' in pattern_type:
                        pattern_color = 'background-color: #dc3545; color: #ffffff;'
                    elif 'Extended Time' in pattern_type:
                        pattern_color = 'background-color: #ffc107; color: #000000;'
                    elif 'High Frequency' in pattern_type:
                        pattern_color = 'background-color: #fd7e14; color: #ffffff;'
                    else:
                        pattern_color = 'background-color: #28a745; color: #ffffff;'
                    
                    roles_display = format_roles_display(pattern.Roles, short=True)
                    
                    print """
                    <tr>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <a href="javascript:viewUser({})" style="color: #17a2b8; font-weight: bold; text-decoration: underline;">
                                {}
                            </a>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            {}
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span style="display: inline-block; padding: 4px 8px; font-size: 11px; font-weight: bold; border-radius: 4px; {}">{}</span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center; font-weight: bold;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{} hours</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{} IPs</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; font-size: 11px;">
                            {} to {}
                        </td>
                    </tr>
                    """.format(
                        pattern.PeopleId, pattern.Username,
                        roles_display,
                        pattern_color, pattern_type,
                        pattern.TotalEvents,
                        pattern.DaysActive,
                        pattern.UniqueHours,
                        pattern.UniqueIPs,
                        format_datetime_short(pattern.FirstEvent),
                        format_datetime_short(pattern.LastEvent)
                    )
                
                print """
            </tbody>
        </table>
        
        <div style="margin-top: 20px; padding: 15px; background-color: #d1ecf1; border-radius: 4px; border-left: 4px solid #17a2b8;">
            <strong><i class="fa fa-info-circle"></i> Pattern Analysis:</strong> 
            <ul style="margin-top: 10px; margin-bottom: 0;">
                <li><strong>Multiple IP Pattern:</strong> User accessed from 5+ different IP addresses (potential compromise)</li>
                <li><strong>Extended Time Pattern:</strong> Login attempts across 12+ different hours (unusual activity)</li>
                <li><strong>High Frequency Pattern:</strong> 10+ failed login attempts (potential brute force target)</li>
                <li><strong>Standard Pattern:</strong> Normal failed login activity within expected parameters</li>
            </ul>
        </div>
    </div>
</div>
                """
                
        except Exception as e:
            print_error("Login Pattern Analysis Generation", e)
    
    def generate_highrisk_ip_analysis(self, lookback_days, monitored_roles):
        """Generate IP analysis specific to high-risk users"""
        try:
            # Build role condition
            if monitored_roles:
                role_conditions = []
                for role in monitored_roles:
                    role_conditions.append("ul.Roles LIKE '%{}%'".format(role))
                role_where = " AND ({})".format(" OR ".join(role_conditions))
            else:
                role_where = " AND (ul.Roles LIKE '%Admin%' OR ul.Roles LIKE '%Finance%' OR ul.Roles LIKE '%Developer%' OR ul.Roles LIKE '%ApiOnly%')"
            
            # ::STEP:: Fixed Query IP Analysis - corrected UserList usage
            sql_ip_analysis = """
                SELECT 
                    al.ClientIp,
                    COUNT(*) AS TotalAttempts,
                    COUNT(DISTINCT ul.UserId) AS HighRiskUsersTargeted,
                    COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) AS DaysActive,
                    MIN(al.ActivityDate) AS FirstSeen,
                    MAX(al.ActivityDate) AS LastSeen,
                    -- Enhanced user status information
                    SUM(CASE WHEN u.IsLockedOut = 1 THEN 1 ELSE 0 END) AS CurrentlyLockedTargets,
                    SUM(CASE WHEN u.LastLockedOutDate >= DATEADD(DAY, -{}, GETDATE()) THEN 1 ELSE 0 END) AS RecentlyLockedTargets,
                    STRING_AGG(ul.Username + '[' + 
                        CASE 
                            WHEN ul.Roles LIKE '%Admin%' THEN 'Admin'
                            WHEN ul.Roles LIKE '%Finance%' THEN 'Finance'
                            WHEN ul.Roles LIKE '%Developer%' THEN 'Dev'
                            ELSE 'Other'
                        END + ']', ', ') AS TargetedUsers
                FROM ActivityLog al
                JOIN UserList ul ON al.UserId = ul.UserId
                LEFT JOIN Users u ON al.UserId = u.UserId
                WHERE al.Activity LIKE '%failed%'
                    AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                    {}
                GROUP BY al.ClientIp
                HAVING COUNT(*) >= 1
                ORDER BY TotalAttempts DESC, HighRiskUsersTargeted DESC;
            """.format(lookback_days, lookback_days, role_where)
            
            ip_data = q.QuerySql(sql_ip_analysis)
            
            if ip_data and len(ip_data) > 0:
                print """
<div style="border: 2px solid #fd7e14; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
    <div style="background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-bottom: 1px solid #ffdf7e; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #333333;">
        <i class="fa fa-crosshairs"></i> IP Addresses Targeting High-Risk Users
    </div>
    <div style="padding: 20px; overflow-x: auto;">
        <table style="width: 100%; border-collapse: collapse; margin: 0; font-size: 13px;">
            <thead>
                <tr>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">IP Address</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Total Attempts</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">High-Risk Users</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Days Active</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">First/Last Seen</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Targeted Users</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left;">Actions</th>
                </tr>
            </thead>
            <tbody>
                """
                
                row_counter = 0
                for ip in ip_data:
                    row_counter += 1
                    
                    # Risk assessment based on attempts and users targeted
                    if ip.TotalAttempts >= 10 or ip.HighRiskUsersTargeted >= 3:
                        row_bg = 'background-color: #f8d7da;'
                        risk_level = 'Critical'
                    elif ip.TotalAttempts >= 5 or ip.HighRiskUsersTargeted >= 2:
                        row_bg = 'background-color: #fff3cd;'
                        risk_level = 'High'
                    else:
                        row_bg = ''
                        risk_level = 'Medium'
                    
                    targeted_users = getattr(ip, 'TargetedUsers', 'Unknown')
                    
                    print """
                    <tr style="{}">
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tphighrisk2025_clickable" 
                                  onclick="openIPLookup('{}')" 
                                  style="font-family: monospace; background-color: #f8f9fa; padding: 4px 8px; border-radius: 4px; font-weight: bold;">
                                {} <i class="fa fa-external-link"></i>
                            </span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">
                            <span class="tphighrisk2025_clickable" onclick="toggleDetails('ip_attempts_{}')">
                                <strong>{}</strong> <i class="fa fa-caret-down"></i>
                            </span>
                            <div id="ip_attempts_{}" class="tphighrisk2025_details">
                                <strong>Attack Analysis for {}:</strong><br>
                                • <strong>Total Attempts:</strong> {}<br>
                                • <strong>High-Risk Users Targeted:</strong> {}<br>
                                • <strong>Attack Span:</strong> {} days<br>
                                • <strong>Risk Level:</strong> {}<br>
                                • <strong>Threat Assessment:</strong> IP specifically targeting privileged accounts
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center; font-weight: bold; color: #dc3545;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; font-size: 11px;">
                            {} to {}
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tphighrisk2025_clickable" onclick="toggleDetails('targeted_users_{}')">
                                View Users <i class="fa fa-caret-down"></i>
                            </span>
                            <div id="targeted_users_{}" class="tphighrisk2025_details">
                                <strong>High-Risk Users Targeted by {}:</strong><br>
                                {}
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">
                            <button onclick="openIPLookup('{}')" style="padding: 3px 8px; font-size: 10px; background-color: #fd7e14; color: white; border: none; border-radius: 3px; cursor: pointer;">
                                <i class="fa fa-search"></i> Lookup
                            </button>
                        </td>
                    </tr>
                    """.format(
                        row_bg,
                        ip.ClientIp, ip.ClientIp,
                        row_counter,
                        ip.TotalAttempts,
                        row_counter,
                        ip.ClientIp,
                        ip.TotalAttempts,
                        ip.HighRiskUsersTargeted,
                        ip.DaysActive,
                        risk_level,
                        ip.HighRiskUsersTargeted,
                        ip.DaysActive,
                        format_datetime_short(ip.FirstSeen),
                        format_datetime_short(ip.LastSeen),
                        row_counter,
                        row_counter,
                        ip.ClientIp,
                        targeted_users if targeted_users else 'No user details available',
                        ip.ClientIp
                    )
                
                print """
            </tbody>
        </table>
        
        <div style="margin-top: 20px; padding: 15px; background-color: #fff3cd; border-radius: 4px; border-left: 4px solid #ffc107;">
            <strong><i class="fa fa-exclamation-triangle"></i> High-Risk IP Analysis:</strong> 
            <ul style="margin-top: 10px; margin-bottom: 0;">
                <li><strong>Targeted Attacks:</strong> These IPs are specifically targeting users with elevated privileges</li>
                <li><strong>Enhanced Threat:</strong> Attacks on Admin, Finance, and Developer accounts pose critical risk</li>
                <li><strong>Immediate Action:</strong> Consider blocking IPs with multiple high-risk user targets</li>
                <li><strong>Monitoring:</strong> Continuous surveillance recommended for all listed IP addresses</li>
            </ul>
        </div>
    </div>
</div>
                """
                
        except Exception as e:
            print_error("High-Risk IP Analysis Generation", e)
    
    def generate_highrisk_recommendations(self, lookback_days, critical_threshold, monitored_roles):
        """Generate security recommendations for high-risk users"""
        try:
            print """
<div style="border: 2px solid #28a745; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
    <div style="background: linear-gradient(135deg, #c3e6cb 0%, #badbcc 100%); border-bottom: 1px solid #a9d1bc; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #333333;">
        <i class="fa fa-shield"></i> High-Risk User Security Recommendations
    </div>
    <div style="padding: 20px;">
        <div class="row">
            <div class="col-md-6">
                <h4 style="color: #155724; margin-bottom: 15px;">Immediate Actions for High-Risk Users</h4>
                <div style="padding: 15px; background-color: #f8d7da; border-left: 4px solid #dc3545; border-radius: 4px; margin-bottom: 15px;">
                    <h5 style="color: #721c24; margin-top: 0;">Critical Priority</h5>
                    <ul style="margin: 0; color: #721c24;">
                        <li>Enable mandatory multi-factor authentication for all privileged accounts</li>
                        <li>Force immediate password reset for accounts with {} failed attempts</li>
                        <li>Implement account lockout after 3 failed attempts for high-risk users</li>
                        <li>Review and audit recent activities for compromised accounts</li>
                        <li>Disable accounts showing signs of compromise until investigation</li>
                    </ul>
                </div>
                
                <div style="padding: 15px; background-color: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;">
                    <h5 style="color: #856404; margin-top: 0;">High Priority</h5>
                    <ul style="margin: 0; color: #856404;">
                        <li>Implement privileged access management (PAM) solution</li>
                        <li>Enable enhanced monitoring for Admin and Finance roles</li>
                        <li>Require additional authentication for sensitive operations</li>
                        <li>Implement time-based access restrictions for privileged accounts</li>
                        <li>Regular access reviews and role validation</li>
                    </ul>
                </div>
            </div>
            
            <div class="col-md-6">
                <h4 style="color: #155724; margin-bottom: 15px;">Long-term Security Enhancements</h4>
                <div style="padding: 15px; background-color: #d4edda; border-left: 4px solid #28a745; border-radius: 4px; margin-bottom: 15px;">
                    <h5 style="color: #155724; margin-top: 0;">Strategic Improvements</h5>
                    <ul style="margin: 0; color: #155724;">
                        <li>Deploy behavioral analytics for anomaly detection</li>
                        <li>Implement zero-trust architecture for privileged access</li>
                        <li>Regular security training specific to high-risk users</li>
                        <li>Automated threat response for privileged account attacks</li>
                        <li>Enhanced logging and monitoring for all privileged activities</li>
                    </ul>
                </div>
                
                <div style="padding: 15px; background-color: #d1ecf1; border-left: 4px solid #17a2b8; border-radius: 4px;">
                    <h5 style="color: #0c5460; margin-top: 0;">Monitoring & Compliance</h5>
                    <ul style="margin: 0; color: #0c5460;">
                        <li>Real-time alerts for all high-risk user security events</li>
                        <li>Weekly security reports for privileged account activity</li>
                        <li>Compliance monitoring for regulatory requirements</li>
                        <li>Integration with SIEM systems for centralized monitoring</li>
                        <li>Regular penetration testing focused on privileged accounts</li>
                    </ul>
                </div>
            </div>
        </div>
        
        <div style="margin-top: 20px; padding: 20px; background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); border-radius: 6px; border-left: 4px solid #2196f3;">
            <h4 style="color: #1565c0; margin-top: 0;">Role-Specific Recommendations</h4>
            <div style="color: #1565c0; font-size: 14px;">
                <p><strong>Admin Accounts:</strong> Highest security level - separate accounts for admin tasks, just-in-time access</p>
                <p><strong>Finance Users:</strong> Enhanced transaction monitoring, dual approval for sensitive operations</p>
                <p><strong>Developer Accounts:</strong> Code signing certificates, restricted production access, secure development practices</p>
                <p><strong>API Access:</strong> Token rotation, rate limiting, API gateway security, comprehensive logging</p>
            </div>
        </div>
        
        <div style="margin-top: 20px; padding: 15px; background-color: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;">
            <h5 style="color: #856404; margin-top: 0;">Implementation Timeline</h5>
            <div style="color: #856404; font-size: 13px;">
                <p><strong>Week 1:</strong> Enable MFA and force password resets for all high-risk accounts</p>
                <p><strong>Week 2-4:</strong> Implement enhanced monitoring and access controls</p>
                <p><strong>Month 2:</strong> Deploy privileged access management and behavioral analytics</p>
                <p><strong>Month 3+:</strong> Continuous improvement and compliance monitoring</p>
            </div>
        </div>
    </div>
</div>
            """.format(critical_threshold)
            
        except Exception as e:
            print_error("High-Risk Recommendations Generation", e)
    
    def generate_report_footer(self):
        """Generate report footer with navigation"""
        print """
                    </div>
                </div>
                <div style="text-align: center; margin: 40px 0; padding: 25px; background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border-radius: 8px;">
                    <a href="javascript:history.back()" style="display: inline-block; padding: 12px 24px; margin: 0 10px; background-color: #6c757d; color: #ffffff; text-decoration: none; border-radius: 6px; font-weight: bold;">
                        <i class="fa fa-arrow-left"></i> Back to Configuration
                    </a>
                    <button onclick="window.print()" style="display: inline-block; padding: 12px 24px; margin: 0 10px; background-color: #17a2b8; color: #ffffff; border: none; border-radius: 6px; font-weight: bold; cursor: pointer;">
                        <i class="fa fa-print"></i> Print Report
                    </button>
                </div>
            </div>
        """

# ::START:: Helper Functions
def show_back_to_hub_button():
    """Show back to hub button"""
    print """
    <div class="unifsec2025_backtohub" style="text-align: center; margin: 20px 0;">
        <form method="post" action="#" style="display: inline;">
            <input type="hidden" name="view" value="">
            <button type="submit" class="unifsec2025_backtohub_btn" style="display: inline-block; padding: 10px 20px; background-color: #6c757d; color: white; text-decoration: none; border-radius: 6px; font-weight: bold; transition: background-color 0.3s ease; border: none; cursor: pointer;">
                <i class="fa fa-arrow-left"></i> Back to Security Hub
            </button>
        </form>
    </div>
    
    <script>
        // Helper function to get the correct form submission URL
        function getPyScriptAddress() {{
            let path = window.location.pathname;
            return path.replace("/PyScript/", "/PyScriptForm/");
        }}
        
        // Set form action to correct URL
        document.querySelector('form').action = getPyScriptAddress();
    </script>
    """

def get_highrisk_params():
    """Get parameters for high-risk user monitor"""
    params = {}
    try:
        if hasattr(model, 'Data'):
            params['lookback_days'] = str(getattr(model.Data, 'lookback_days', '30'))
            params['max_results'] = str(getattr(model.Data, 'max_results', '100'))
            params['critical_threshold'] = str(getattr(model.Data, 'critical_threshold', '5'))
            params['high_threshold'] = str(getattr(model.Data, 'high_threshold', '3'))
            params['monitor_admin'] = str(getattr(model.Data, 'monitor_admin', ''))
            params['monitor_finance'] = str(getattr(model.Data, 'monitor_finance', ''))
            params['monitor_financeadmin'] = str(getattr(model.Data, 'monitor_financeadmin', ''))
            params['monitor_developer'] = str(getattr(model.Data, 'monitor_developer', ''))
            params['monitor_apionly'] = str(getattr(model.Data, 'monitor_apionly', ''))
            params['custom_roles'] = str(getattr(model.Data, 'custom_roles', ''))
            params['show_alerts'] = str(getattr(model.Data, 'show_alerts', ''))
            params['show_user_overview'] = str(getattr(model.Data, 'show_user_overview', ''))
            params['show_security_events'] = str(getattr(model.Data, 'show_security_events', ''))
            params['show_role_analysis'] = str(getattr(model.Data, 'show_role_analysis', ''))
            params['show_login_patterns'] = str(getattr(model.Data, 'show_login_patterns', ''))
            params['show_ip_analysis'] = str(getattr(model.Data, 'show_ip_analysis', ''))
            params['show_recommendations'] = str(getattr(model.Data, 'show_recommendations', ''))
    except:
        pass
    return params

def get_enhanced_params():
    """Get parameters for enhanced security dashboard"""
    params = {}
    try:
        if hasattr(model, 'Data'):
            # Basic parameters
            params['lookback_days'] = str(getattr(model.Data, 'lookback_days', '30'))
            params['min_attempts'] = str(getattr(model.Data, 'min_attempts', '5'))
            params['max_results'] = str(getattr(model.Data, 'max_results', '500'))
            params['activity_filter'] = str(getattr(model.Data, 'activity_filter', '%failed%'))
            params['internal_ips'] = str(getattr(model.Data, 'internal_ips', ''))
            
            # Threat detection settings
            params['high_risk_threshold'] = str(getattr(model.Data, 'high_risk_threshold', '20'))
            params['auto_block_threshold'] = str(getattr(model.Data, 'auto_block_threshold', '25'))
            params['time_window'] = str(getattr(model.Data, 'time_window', '30'))
            params['recent_events_limit'] = str(getattr(model.Data, 'recent_events_limit', '50'))
            
            # Report sections
            params['show_critical_alerts'] = str(getattr(model.Data, 'show_critical_alerts', ''))
            params['show_summary'] = str(getattr(model.Data, 'show_summary', ''))
            params['show_recent_activity'] = str(getattr(model.Data, 'show_recent_activity', ''))
            params['show_threat_analysis'] = str(getattr(model.Data, 'show_threat_analysis', ''))
            params['show_detailed'] = str(getattr(model.Data, 'show_detailed', ''))
            params['show_ip_analysis'] = str(getattr(model.Data, 'show_ip_analysis', ''))
            params['show_user_analysis'] = str(getattr(model.Data, 'show_user_analysis', ''))
            params['show_compliance_report'] = str(getattr(model.Data, 'show_compliance_report', ''))
            params['show_successful_logins'] = str(getattr(model.Data, 'show_successful_logins', ''))
            params['exclude_internal'] = str(getattr(model.Data, 'exclude_internal', ''))
            params['show_user_details'] = str(getattr(model.Data, 'show_user_details', ''))
            params['include_remediation'] = str(getattr(model.Data, 'include_remediation', ''))
            params['generate_csv_export'] = str(getattr(model.Data, 'generate_csv_export', ''))
            
            # Advanced options
            params['enable_pattern_analysis'] = str(getattr(model.Data, 'enable_pattern_analysis', ''))
            params['enable_geo_tracking'] = str(getattr(model.Data, 'enable_geo_tracking', ''))
            params['enable_auto_blocking'] = str(getattr(model.Data, 'enable_auto_blocking', ''))
            params['enable_realtime_monitoring'] = str(getattr(model.Data, 'enable_realtime_monitoring', ''))
    except:
        pass
    return params

def format_roles_display(roles_string, short=False):
    """Format roles string for display with highlighting"""
    if not roles_string:
        return '<span style="color: #6c757d;">No roles</span>'
    
    high_risk_roles = ['Admin', 'Finance', 'FinanceAdmin', 'Developer', 'ApiOnly']
    roles = [role.strip() for role in str(roles_string).split('|') if role.strip()]
    formatted_roles = []
    
    for role in roles:
        if role in high_risk_roles:
            if short:
                formatted_roles.append('<span class="tphighrisk2025_rolehigh">{}</span>'.format(role))
            else:
                formatted_roles.append('<span style="background-color: #dc3545; color: #ffffff; padding: 2px 6px; border-radius: 3px; font-size: 11px; font-weight: bold; margin: 1px;">{}</span>'.format(role))
        else:
            if short:
                formatted_roles.append('<span class="tphighrisk2025_rolelow">{}</span>'.format(role))
            else:
                formatted_roles.append('<span style="background-color: #6c757d; color: #ffffff; padding: 2px 6px; border-radius: 3px; font-size: 11px; margin: 1px;">{}</span>'.format(role))
    
    return ' '.join(formatted_roles) if formatted_roles else '<span style="color: #6c757d;">No roles</span>'

def format_datetime(dt):
    """Format .NET DateTime object to string"""
    try:
        if dt is None:
            return 'N/A'
        return str(dt)[:19]
    except:
        return 'N/A'

def format_datetime_short(dt):
    """Format .NET DateTime object to short string"""
    try:
        if dt is None:
            return 'N/A'
        dt_str = str(dt)
        return dt_str[:16]
    except:
        return 'N/A'

def format_time(dt):
    """Format .NET DateTime object to time only"""
    try:
        if dt is None:
            return 'N/A'
        dt_str = str(dt)
        if ' ' in dt_str:
            return dt_str.split(' ')[1][:8]
        return dt_str[:8]
    except:
        return 'N/A'

def print_error(operation, exception):
    """Print formatted error message with expandable details"""
    import random
    error_id = "error_{}".format(random.randint(1000, 9999))
    
    print """
    <div class="alert alert-danger" style="margin: 20px 0;">
        <h4 style="margin-top: 0;">
            <i class="fa fa-exclamation-circle"></i> Error in {}
        </h4>
        <p><strong>Error Message:</strong> {}</p>
        <p>
            <a href="javascript:void(0);" onclick="toggleDetails('{}')" style="color: #721c24; text-decoration: underline;">
                <i class="fa fa-chevron-down"></i> Click here for technical details
            </a>
        </p>
        <div id="{}" class="tphighrisk2025_details" style="margin-top: 15px; background-color: #f5c6cb; padding: 15px; border-radius: 4px; display: none;">
            <strong>Stack Trace:</strong>
            <pre style="background-color: #ffffff; padding: 10px; border: 1px solid #f5c6cb; border-radius: 4px; overflow-x: auto;">""".format(
        operation, 
        str(exception), 
        error_id,
        error_id
    )
    
    traceback.print_exc()
    
    print """</pre>
        </div>
    </div>
    """

def print_critical_error(operation, exception):
    """Print critical error message"""
    print "<h2>Critical Error</h2>"
    print "<div class='alert alert-danger' style='padding: 20px; margin: 20px; border: 2px solid #dc3545; border-radius: 8px;'>"
    print "<p><strong>Error in {}:</strong> {}</p>".format(operation, str(exception))
    print "<pre style='background-color: #f8f9fa; padding: 15px; border-radius: 4px; overflow-x: auto;'>"
    traceback.print_exc()
    print "</pre>"
    print "</div>"
    
class EnhancedSecurityDashboard:
    """Advanced security monitoring and threat detection system"""
    
    def __init__(self):
        # ::START:: Enhanced Configuration Parameters
        self.lookback_days = 30
        self.high_risk_threshold = 20  # Increased from 10
        self.medium_risk_threshold = 10  # Increased from 5
        self.brute_force_time_window = 30  # Tightened to 30 minutes
        self.max_display_limit = 500  # Increased limit from 10
        self.recent_events_limit = 50  # For recent activity monitoring
        self.geo_tracking_enabled = True
        self.auto_block_threshold = 25  # Automatic blocking recommendation
        
    def get_configuration_form(self):
        """Generate enhanced configuration form with advanced security options"""
        try:
            # ::START:: Enhanced Configuration Form Generation
            print """
            <style>
            .tpsecuritydash2025_configcontainer {{
                max-width: 1400px;
                margin: 0 auto;
                padding: 20px;
                font-family: Arial, sans-serif;
            }}
            .tpsecuritydash2025_configpanel {{
                border: 2px solid #6c757d;
                border-radius: 8px;
                margin: 20px 0;
                background-color: #ffffff;
                box-shadow: 0 4px 6px rgba(0,0,0,0.15);
            }}
            .tpsecuritydash2025_configheader {{
                background: linear-gradient(135deg, #e9ecef 0%, #dee2e6 100%);
                border-bottom: 1px solid #adb5bd;
                padding: 15px 20px;
                border-radius: 6px 6px 0 0;
                font-weight: bold;
                font-size: 18px;
                color: #333;
            }}
            .tpsecuritydash2025_configbody {{
                padding: 25px;
            }}
            .tpsecuritydash2025_formrow {{
                margin-bottom: 20px;
                display: block;
                width: 100%;
            }}
            .tpsecuritydash2025_formcol {{
                display: inline-block;
                vertical-align: top;
                width: 48%;
                margin-right: 2%;
                margin-bottom: 15px;
            }}
            .tpsecuritydash2025_formcolthird {{
                display: inline-block;
                vertical-align: top;
                width: 31%;
                margin-right: 2%;
                margin-bottom: 15px;
            }}
            .tpsecuritydash2025_formgroup {{
                margin-bottom: 20px;
            }}
            .tpsecuritydash2025_formlabel {{
                display: block;
                margin-bottom: 8px;
                font-weight: bold;
                color: #495057;
            }}
            .tpsecuritydash2025_formcontrol {{
                width: 100%;
                padding: 10px 12px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-size: 14px;
                background-color: #fff;
                box-sizing: border-box;
            }}
            .tpsecuritydash2025_formcontrol:focus {{
                outline: 0;
                border-color: #80bdff;
                box-shadow: 0 0 0 0.2rem rgba(0,123,255,.25);
            }}
            .tpsecuritydash2025_checkbox {{
                margin-bottom: 12px;
                display: flex;
                align-items: center;
            }}
            .tpsecuritydash2025_checkbox input {{
                margin-right: 8px;
            }}
            .tpsecuritydash2025_btncontainer {{
                text-align: center;
                margin-top: 30px;
            }}
            .tpsecuritydash2025_btn {{
                display: inline-block;
                padding: 12px 24px;
                font-size: 16px;
                font-weight: bold;
                text-decoration: none;
                border: none;
                border-radius: 6px;
                cursor: pointer;
                background-color: #007bff;
                color: #ffffff;
                transition: background-color 0.3s;
                margin: 0 10px;
            }}
            .tpsecuritydash2025_btn:hover {{
                background-color: #0056b3;
            }}
            .tpsecuritydash2025_btn_danger {{
                background-color: #dc3545;
            }}
            .tpsecuritydash2025_btn_danger:hover {{
                background-color: #c82333;
            }}
            .tpsecuritydash2025_loading {{
                display: none;
                margin-top: 15px;
                color: #6c757d;
            }}
            .tpsecuritydash2025_helptext {{
                font-size: 12px;
                color: #6c757d;
                margin-top: 5px;
            }}
            .tpsecuritydash2025_alert {{
                padding: 15px;
                margin-bottom: 20px;
                border: 1px solid transparent;
                border-radius: 4px;
                background-color: #cce7ff;
                border-color: #b3d9ff;
                color: #004085;
            }}
            .tpsecuritydash2025_alert_warning {{
                background-color: #fff3cd;
                border-color: #ffeaa7;
                color: #856404;
            }}
            .tpsecuritydash2025_advanced {{
                display: none;
                margin-top: 15px;
                padding: 15px;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
            }}
            </style>
            
            <div class="tpsecuritydash2025_configcontainer">
                <h2><i class="fa fa-shield"></i> Enhanced Security Dashboard Configuration</h2>
                <div class="tpsecuritydash2025_alert">
                    <strong>Enhanced Security Monitoring:</strong> This advanced dashboard provides comprehensive threat detection, 
                    real-time monitoring, and automated security analysis for your TouchPoint system.
                </div>
                
                <form method="post" action="#">
                    <input type="hidden" name="view" value="enhanced">
                    <input type="hidden" name="action" value="generate_report">
                    
                    <div class="tpsecuritydash2025_configpanel">
                        <div class="tpsecuritydash2025_configheader">
                            <i class="fa fa-search"></i> Analysis Parameters
                        </div>
                        <div class="tpsecuritydash2025_configbody">
                            <div class="tpsecuritydash2025_formrow">
                                <div class="tpsecuritydash2025_formcolthird">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel" for="lookback_days">Time Period:</label>
                                        <select class="tpsecuritydash2025_formcontrol" name="lookback_days" id="lookback_days">
                                            <option value="1">Last 24 Hours</option>
                                            <option value="3">Last 3 Days</option>
                                            <option value="7">Last 7 Days</option>
                                            <option value="14">Last 2 Weeks</option>
                                            <option value="30" selected>Last 30 Days</option>
                                            <option value="60">Last 2 Months</option>
                                            <option value="90">Last 3 Months</option>
                                            <option value="180">Last 6 Months</option>
                                        </select>
                                    </div>
                                </div>
                                
                                <div class="tpsecuritydash2025_formcolthird">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel" for="min_attempts">Minimum Attempts:</label>
                                        <select class="tpsecuritydash2025_formcontrol" name="min_attempts" id="min_attempts">
                                            <option value="1">1+ Attempts</option>
                                            <option value="3">3+ Attempts</option>
                                            <option value="5" selected>5+ Attempts</option>
                                            <option value="10">10+ Attempts</option>
                                            <option value="15">15+ Attempts</option>
                                            <option value="25">25+ Attempts</option>
                                            <option value="50">50+ Attempts</option>
                                        </select>
                                    </div>
                                </div>
                                
                                <div class="tpsecuritydash2025_formcolthird">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel" for="max_results">Maximum Results:</label>
                                        <select class="tpsecuritydash2025_formcontrol" name="max_results" id="max_results">
                                            <option value="50">Top 50 Results</option>
                                            <option value="100">Top 100 Results</option>
                                            <option value="250">Top 250 Results</option>
                                            <option value="500" selected>Top 500 Results</option>
                                            <option value="1000">Top 1000 Results</option>
                                            <option value="0">All Results (No Limit)</option>
                                        </select>
                                        <div class="tpsecuritydash2025_helptext">Limit results for performance</div>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="tpsecuritydash2025_formrow">
                                <div class="tpsecuritydash2025_formcol">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel" for="activity_filter">Security Event Filter:</label>
                                        <select class="tpsecuritydash2025_formcontrol" name="activity_filter" id="activity_filter">
                                            <option value="%failed%" selected>All Failed Activities</option>
                                            <option value="%password%">Password Related Events</option>
                                            <option value="%logged%">Failed Login Attempts</option>
                                            <option value="%locked%">Account Lockouts</option>
                                            <option value="ForgotPassword">Password Reset Attempts</option>
                                            <option value="User account locked">Account Locked Events</option>
                                            <option value="%invalid%">Invalid Login Attempts</option>
                                            <option value="%mobile%">Mobile App Failures</option>
                                        </select>
                                    </div>
                                </div>
                                
                                <div class="tpsecuritydash2025_formcol">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel" for="internal_ips">Internal IP Ranges:</label>
                                        <input type="text" class="tpsecuritydash2025_formcontrol" name="internal_ips" 
                                               id="internal_ips" placeholder="e.g., 192.168.1.,10.0.0.,172.16.">
                                        <div class="tpsecuritydash2025_helptext">Comma-separated internal IP prefixes</div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="tpsecuritydash2025_configpanel">
                        <div class="tpsecuritydash2025_configheader">
                            <i class="fa fa-exclamation-triangle"></i> Threat Detection Settings
                        </div>
                        <div class="tpsecuritydash2025_configbody">
                            <div class="tpsecuritydash2025_formrow">
                                <div class="tpsecuritydash2025_formcolthird">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel" for="high_risk_threshold">High Risk Threshold:</label>
                                        <input type="number" class="tpsecuritydash2025_formcontrol" name="high_risk_threshold" 
                                               id="high_risk_threshold" value="20" min="1" max="1000">
                                        <div class="tpsecuritydash2025_helptext">Failed attempts for high risk alert</div>
                                    </div>
                                </div>
                                
                                <div class="tpsecuritydash2025_formcolthird">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel" for="auto_block_threshold">Auto-Block Threshold:</label>
                                        <input type="number" class="tpsecuritydash2025_formcontrol" name="auto_block_threshold" 
                                               id="auto_block_threshold" value="25" min="1" max="1000">
                                        <div class="tpsecuritydash2025_helptext">Attempts to recommend IP blocking</div>
                                    </div>
                                </div>
                                
                                <div class="tpsecuritydash2025_formcolthird">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel" for="time_window">Brute Force Window (Minutes):</label>
                                        <input type="number" class="tpsecuritydash2025_formcontrol" name="time_window" 
                                               id="time_window" value="30" min="1" max="1440">
                                        <div class="tpsecuritydash2025_helptext">Time window for attack detection</div>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="tpsecuritydash2025_formrow">
                                <div class="tpsecuritydash2025_formcol">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel" for="recent_events_limit">Recent Events Limit:</label>
                                        <select class="tpsecuritydash2025_formcontrol" name="recent_events_limit" id="recent_events_limit">
                                            <option value="25">Last 25 Events</option>
                                            <option value="50" selected>Last 50 Events</option>
                                            <option value="100">Last 100 Events</option>
                                            <option value="250">Last 250 Events</option>
                                            <option value="500">Last 500 Events</option>
                                        </select>
                                        <div class="tpsecuritydash2025_helptext">Number of recent events to display in timeline</div>
                                    </div>
                                </div>
                                
                                <div class="tpsecuritydash2025_formcol">
                                    <div class="tpsecuritydash2025_formgroup">
                                        <label class="tpsecuritydash2025_formlabel">Advanced Detection:</label>
                                        <div class="tpsecuritydash2025_checkbox">
                                            <input type="checkbox" name="enable_pattern_analysis" value="1" checked>
                                            <label>Enable Pattern Analysis</label>
                                        </div>
                                        <div class="tpsecuritydash2025_checkbox">
                                            <input type="checkbox" name="enable_geo_tracking" value="1" checked>
                                            <label>Enable Geographic Tracking</label>
                                        </div>
                                        <div class="tpsecuritydash2025_checkbox">
                                            <input type="checkbox" name="enable_auto_blocking" value="1">
                                            <label>Show Auto-Block Recommendations</label>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="tpsecuritydash2025_configpanel">
                        <div class="tpsecuritydash2025_configheader">
                            <i class="fa fa-cog"></i> Report Sections
                        </div>
                        <div class="tpsecuritydash2025_configbody">
                            <div class="tpsecuritydash2025_formrow">
                                <div class="tpsecuritydash2025_formcol">
                                    <label class="tpsecuritydash2025_formlabel">Primary Reports:</label>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_critical_alerts" value="1" checked>
                                        <label><strong>Critical Security Alerts</strong></label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_summary" value="1" checked>
                                        <label><strong>Executive Summary & KPIs</strong></label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_recent_activity" value="1" checked>
                                        <label><strong>Recent Security Events Timeline</strong></label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_threat_analysis" value="1" checked>
                                        <label><strong>Advanced Threat Analysis</strong></label>
                                    </div>
                                </div>
                                
                                <div class="tpsecuritydash2025_formcol">
                                    <label class="tpsecuritydash2025_formlabel">Detailed Analysis:</label>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_detailed" value="1" checked>
                                        <label>Detailed Security Events</label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_ip_analysis" value="1" checked>
                                        <label>IP Address Analysis</label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_user_analysis" value="1">
                                        <label>User Behavior Analysis</label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_compliance_report" value="1">
                                        <label>Compliance & Audit Report</label>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="tpsecuritydash2025_formrow">
                                <div class="tpsecuritydash2025_formcol">
                                    <label class="tpsecuritydash2025_formlabel">Additional Options:</label>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_successful_logins" value="1">
                                        <label>Include Successful Login Comparison</label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="exclude_internal" value="1">
                                        <label>Exclude Internal IP Addresses</label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="show_user_details" value="1">
                                        <label>Show User Details (Expandable)</label>
                                    </div>
                                </div>
                                
                                <div class="tpsecuritydash2025_formcol">
                                    <label class="tpsecuritydash2025_formlabel">Export & Automation:</label>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="include_remediation" value="1" checked>
                                        <label>Include Remediation Recommendations</label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="generate_csv_export" value="1">
                                        <label>Generate CSV Export Data</label>
                                    </div>
                                    <div class="tpsecuritydash2025_checkbox">
                                        <input type="checkbox" name="enable_realtime_monitoring" value="1">
                                        <label>Enable Real-time Monitoring Mode</label>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="tpsecuritydash2025_btncontainer">
                        <button type="submit" class="tpsecuritydash2025_btn">
                            <i class="fa fa-search"></i> Generate Advanced Security Report
                        </button>
                        <span id="tpsecuritydash2025_loading" class="tpsecuritydash2025_loading" style="display: none; margin-left: 15px;">
                            <i class="fa fa-spinner fa-spin"></i> Analyzing security data...
                        </span>
                    </div>
                </form>
                
                <div class="tpsecuritydash2025_alert tpsecuritydash2025_alert_warning" style="margin-top: 20px;">
                    <strong>Security Note:</strong> This enhanced dashboard can process large amounts of data. 
                    For performance, consider limiting results when analyzing extended time periods or enable 
                    real-time monitoring for continuous security oversight.
                </div>
            </div>
            
            <script>
                // Helper function to get the correct form submission URL
                function getPyScriptAddress() {{
                    let path = window.location.pathname;
                    return path.replace("/PyScript/", "/PyScriptForm/");
                }}
                
                // Set form action to correct URL
                document.querySelector('form').action = getPyScriptAddress();
                
                // Show loading indicator on form submit
                document.querySelector('form').addEventListener('submit', function(e) {{
                    document.getElementById('tpsecuritydash2025_loading').style.display = 'inline-block';
                    document.querySelector('button[type="submit"]').disabled = true;
                }});
                
                // Emergency report function
                function generateEmergencyReport() {{
                    if (confirm('Generate an emergency threat scan for the last 24 hours?')) {{
                        // Set emergency parameters
                        document.getElementById('lookback_days').value = '1';
                        document.getElementById('min_attempts').value = '3';
                        document.getElementById('high_risk_threshold').value = '10';
                        document.getElementById('time_window').value = '15';
                        
                        // Enable critical sections
                        document.querySelector('input[name="show_critical_alerts"]').checked = true;
                        document.querySelector('input[name="show_recent_activity"]').checked = true;
                        document.querySelector('input[name="show_threat_analysis"]').checked = true;
                        
                        // Submit form
                        document.querySelector('form').submit();
                    }}
                }}
            </script>
            """
            
        except Exception as e:
            self.print_error("Configuration Form Generation", e)
            
    def generate_role_based_analysis(self, lookback_days, monitored_roles):
        """Generate analysis by role type"""
        try:
            print """
    <div style="border: 2px solid #6f42c1; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
        <div style="background: linear-gradient(135deg, #e2d9f3 0%, #d6c7f0 100%); border-bottom: 1px solid #c7b3ed; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #333333;">
            <i class="fa fa-users-cog"></i> Role-Based Security Risk Analysis (Last {} Days)
        </div>
        <div style="padding: 20px;">
            """.format(lookback_days)
            
            # ::STEP:: OPTIMIZED - Single query for all roles to avoid multiple table scans
            if monitored_roles:
                # Build role conditions for a single query
                role_conditions = []
                for role in monitored_roles:
                    role_conditions.append("ul.Roles LIKE '%{}%'".format(role))
                role_where = " OR ".join(role_conditions)
                
                sql_all_roles = """
                    SELECT 
                        -- Extract role name for grouping
                        CASE 
                            WHEN ul.Roles LIKE '%Admin%' THEN 'Admin'
                            WHEN ul.Roles LIKE '%Finance%' THEN 'Finance' 
                            WHEN ul.Roles LIKE '%FinanceAdmin%' THEN 'FinanceAdmin'
                            WHEN ul.Roles LIKE '%Developer%' THEN 'Developer'
                            WHEN ul.Roles LIKE '%ApiOnly%' THEN 'ApiOnly'
                            ELSE 'Other'
                        END AS RoleName,
                        
                        -- Count unique users with this role (total)
                        COUNT(DISTINCT ul.UserId) AS TotalUsersWithRole,
                        
                        -- Count users with failures in the period
                        COUNT(DISTINCT CASE WHEN al.ActivityDate IS NOT NULL THEN ul.UserId END) AS UsersWithFailures,
                        
                        -- Count total failures in the period
                        COUNT(al.ActivityDate) AS TotalFailures,
                        
                        -- Count unique IPs attacking users with this role
                        COUNT(DISTINCT al.ClientIp) AS UniqueIPs,
                        
                        -- Count unique days with incidents
                        COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) AS DaysWithIncidents,
                        
                        -- Current account status (not time-limited)
                        COUNT(DISTINCT CASE WHEN u.IsLockedOut = 1 THEN ul.UserId END) AS CurrentlyLockedUsers,
                        COUNT(DISTINCT CASE WHEN u.LastLockedOutDate >= DATEADD(DAY, -{}, GETDATE()) THEN ul.UserId END) AS RecentlyLockedUsers,
                        COUNT(DISTINCT CASE WHEN u.MustChangePassword = 1 THEN ul.UserId END) AS ForcedPasswordResets
                        
                    FROM UserList ul
                    LEFT JOIN ActivityLog al ON ul.UserId = al.UserId 
                        AND al.Activity LIKE '%failed%'
                        AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                    LEFT JOIN Users u ON ul.UserId = u.UserId
                    WHERE ({})
                    GROUP BY 
                        CASE 
                            WHEN ul.Roles LIKE '%Admin%' THEN 'Admin'
                            WHEN ul.Roles LIKE '%Finance%' THEN 'Finance'
                            WHEN ul.Roles LIKE '%FinanceAdmin%' THEN 'FinanceAdmin' 
                            WHEN ul.Roles LIKE '%Developer%' THEN 'Developer'
                            WHEN ul.Roles LIKE '%ApiOnly%' THEN 'ApiOnly'
                            ELSE 'Other'
                        END
                    ORDER BY TotalFailures DESC;
                """.format(lookback_days, lookback_days, role_where)
                
                try:
                    role_data = q.QuerySql(sql_all_roles)
                    if role_data and len(role_data) > 0:
                        for role_stats in role_data:
                            role_name = getattr(role_stats, 'RoleName', 'Unknown')
                            
                            # Skip if this role wasn't in our monitored list
                            if role_name not in monitored_roles and role_name != 'Other':
                                continue
                                
                            self.display_role_analysis_results(role_name, role_stats, lookback_days)
                    else:
                        print """
                        <div style="margin: 15px 0; padding: 15px; border: 1px solid #28a745; border-radius: 6px; background-color: #d4edda; color: #155724;">
                            <strong>No Security Incidents:</strong> No users with monitored roles experienced security incidents in the last {} days.
                        </div>
                        """.format(lookback_days)
                except Exception as role_error:
                    print """
                    <div style="margin: 15px 0; padding: 15px; border: 1px solid #dc3545; border-radius: 6px; background-color: #f8d7da;">
                        <h4 style="color: #721c24; margin-top: 0;">
                            Role Analysis Error
                        </h4>
                        <p style="color: #721c24; margin: 0;">
                            Error analyzing roles: {}
                        </p>
                    </div>
                    """.format(str(role_error)[:100])
            else:
                print """
                <div style="margin: 15px 0; padding: 15px; border: 1px solid #ffc107; border-radius: 6px; background-color: #fff3cd;">
                    <strong>No Roles Selected:</strong> No high-risk roles selected for monitoring.
                </div>
                """
        
        except Exception as e:
            self.print_error("Configuration Form Generation", e)
                
    def generate_security_report(self, params):
        """Generate comprehensive enhanced security report"""
        try:
            # ::START:: Enhanced Parameter Processing
            lookback_days = int(params.get('lookback_days', 30))
            min_attempts = int(params.get('min_attempts', 5))
            max_results = int(params.get('max_results', 500))
            high_risk_threshold = int(params.get('high_risk_threshold', 20))
            auto_block_threshold = int(params.get('auto_block_threshold', 25))
            time_window = int(params.get('time_window', 30))
            recent_events_limit = int(params.get('recent_events_limit', 50))
            
            activity_filter = params.get('activity_filter', '%failed%')
            internal_ips = params.get('internal_ips', '')
            
            # Enhanced boolean parameters
            show_critical_alerts = params.get('show_critical_alerts') == '1'
            show_summary = params.get('show_summary') == '1'
            show_recent_activity = params.get('show_recent_activity') == '1'
            show_threat_analysis = params.get('show_threat_analysis') == '1'
            show_detailed = params.get('show_detailed') == '1'
            show_ip_analysis = params.get('show_ip_analysis') == '1'
            show_user_analysis = params.get('show_user_analysis') == '1'
            show_compliance_report = params.get('show_compliance_report') == '1'
            show_successful_logins = params.get('show_successful_logins') == '1'
            exclude_internal = params.get('exclude_internal') == '1'
            show_user_details = params.get('show_user_details') == '1'
            include_remediation = params.get('include_remediation') == '1'
            enable_pattern_analysis = params.get('enable_pattern_analysis') == '1'
            enable_geo_tracking = params.get('enable_geo_tracking') == '1'
            enable_auto_blocking = params.get('enable_auto_blocking') == '1'
            enable_realtime_monitoring = params.get('enable_realtime_monitoring') == '1'
            
            activity_filter_display = self.get_activity_filter_display_name(activity_filter)
            
            # ::START:: Report Header with Enhanced Styling
            print """
            <div class="container-fluid">
                <div class="row">
                    <div class="col-md-12">
                        <div style="border-bottom: 3px solid #dc3545; padding-bottom: 20px; margin-bottom: 30px; background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); padding: 20px; border-radius: 8px;">
                            <h1 style="color: #dc3545; font-size: 32px; margin: 0; font-weight: bold;">
                                <i class="fa fa-shield"></i> Enhanced Security Analysis Report
                            </h1>
                            <p style="color: #6c757d; margin: 15px 0 0 0; font-size: 16px;">
                                <strong>Analysis Period:</strong> Last {} days | 
                                <strong>Filter:</strong> {} | 
                                <strong>Min Attempts:</strong> {} | 
                                <strong>Max Results:</strong> {} | 
                                <strong>Generated:</strong> {}
                            </p>
                            {}
                        </div>
                        
                        <style type="text/css">
                        /* Enhanced TouchPoint Security Dashboard 2025 - Report Styles */
                        .tpsecuritydash2025_reportcontainer {{
                            font-family: Arial, sans-serif !important;
                            max-width: 1600px !important;
                            margin: 0 auto !important;
                        }}
                        .tpsecuritydash2025_section {{
                            margin-bottom: 45px !important;
                        }}
                        .tpsecuritydash2025_backbuttons {{
                            margin: 40px 0 !important;
                            text-align: center !important;
                            padding: 25px !important;
                            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%) !important;
                            border-radius: 8px !important;
                            border: 1px solid #dee2e6 !important;
                        }}
                        .tpsecuritydash2025_backbtn {{
                            display: inline-block !important;
                            padding: 12px 24px !important;
                            margin: 0 10px !important;
                            background-color: #6c757d !important;
                            color: #ffffff !important;
                            text-decoration: none !important;
                            border-radius: 6px !important;
                            font-weight: bold !important;
                            border: none !important;
                            cursor: pointer !important;
                            font-size: 14px !important;
                        }}
                        .tpsecuritydash2025_backbtn:hover {{
                            background-color: #5a6268 !important;
                            color: #ffffff !important;
                            text-decoration: none !important;
                        }}
                        .tpsecuritydash2025_printbtn {{
                            background-color: #17a2b8 !important;
                        }}
                        .tpsecuritydash2025_printbtn:hover {{
                            background-color: #138496 !important;
                        }}
                        .tpsecuritydash2025_exportbtn {{
                            background-color: #28a745 !important;
                        }}
                        .tpsecuritydash2025_exportbtn:hover {{
                            background-color: #218838 !important;
                        }}
                        .tpsecuritydash2025_criticalAlert {{
                            background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%) !important;
                            border: 3px solid #dc3545 !important;
                            padding: 25px !important;
                            margin: 25px 0 !important;
                            border-radius: 10px !important;
                            font-family: Arial, sans-serif !important;
                            color: #721c24 !important;
                            box-shadow: 0 4px 8px rgba(220, 53, 69, 0.2) !important;
                        }}
                        .tpsecuritydash2025_emergencyAlert {{
                            background: linear-gradient(135deg, #ff6b6b 0%, #dc3545 100%) !important;
                            border: 3px solid #c82333 !important;
                            color: #ffffff !important;
                            animation: pulse 2s infinite !important;
                        }}
                        @keyframes pulse {{
                            0% {{ box-shadow: 0 0 0 0 rgba(220, 53, 69, 0.7); }}
                            70% {{ box-shadow: 0 0 0 10px rgba(220, 53, 69, 0); }}
                            100% {{ box-shadow: 0 0 0 0 rgba(220, 53, 69, 0); }}
                        }}
                        .tpsecuritydash2025_clickable {{
                            color: #ffffff !important; /* White text on blue background */
                            background-color: #007bff !important;
                            padding: 6px 12px !important;
                            border-radius: 4px !important;
                            display: inline-block !important;
                            font-weight: bold !important;
                            text-decoration: none !important;
                            cursor: pointer !important;
                            border: 1px solid #0056b3 !important;
                        }}
                        .tpsecuritydash2025_clickable:hover {{
                            background-color: #0056b3 !important;
                            color: #ffffff !important;
                            border-color: #004085 !important;
                            text-decoration: none !important;
                        }}
                        .tpsecuritydash2025_details {{
                            display: none;
                            margin-top: 15px;
                            padding: 15px;
                            background-color: #f8f9fa;
                            border: 1px solid #dee2e6;
                            border-radius: 6px;
                            border-left: 4px solid #007bff;
                        }}
                        .tpsecuritydash2025_clickable strong {{
                            color: #ffffff !important;
                        }}
                        .tpsecuritydash2025_realtimebadge {{
                            background-color: #28a745 !important;
                            color: #ffffff !important;
                            padding: 4px 8px !important;
                            border-radius: 4px !important;
                            font-size: 12px !important;
                            font-weight: bold !important;
                            animation: blink 1.5s linear infinite !important;
                        }}
                        @keyframes blink {{
                            0%, 50% {{ opacity: 1; }}
                            51%, 100% {{ opacity: 0.5; }}
                        }}
                        </style>
                        
                        <script>
                        function toggleDetails(id) {{
                            var element = document.getElementById(id);
                            if (element.style.display === "none" || element.style.display === "") {{
                                element.style.display = "block";
                            }} else {{
                                element.style.display = "none";
                            }}
                        }}
                        
                        function scrollToSection(sectionId) {{
                            var section = document.getElementById(sectionId);
                            if (section) {{
                                section.scrollIntoView({{ behavior: 'smooth' }});
                                // Highlight briefly
                                section.style.backgroundColor = '#ffffcc';
                                setTimeout(function() {{
                                    section.style.backgroundColor = '';
                                }}, 3000);
                            }}
                        }}
                        
                        function openIPLookup(ip) {{
                            if (ip) {{
                                var url = 'https://www.abuseipdb.com/check/' + ip;
                                window.open(url, '_blank');
                            }}
                        }}
                        
                        function viewUser(userId) {{
                            if (userId && userId !== 'NULL' && userId !== '0') {{
                                window.open('/Person2/' + userId, '_blank');
                            }} else {{
                                alert('User ID not available');
                            }}
                        }}
                        
                        function exportCSV() {{
                            alert('CSV export functionality would be implemented here');
                        }}
                        
                        function refreshRealtime() {{
                            if (confirm('Refresh real-time data? This will reload the page.')) {{
                                window.location.reload();
                            }}
                        }}
                        
                        // Auto-refresh for real-time monitoring
                        {}
                        </script>
            """.format(
                lookback_days, 
                activity_filter_display, 
                min_attempts,
                max_results if max_results > 0 else "Unlimited",
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                '<span class="tpsecuritydash2025_realtimebadge"><i class="fa fa-refresh"></i> REAL-TIME MODE</span>' if enable_realtime_monitoring else '',
                'setInterval(function() { document.getElementById("realtime-refresh").style.opacity = document.getElementById("realtime-refresh").style.opacity == "0.5" ? "1" : "0.5"; }, 1000);' if enable_realtime_monitoring else ''
            )
            
            # ::STEP:: Generate Critical Security Alerts
            if show_critical_alerts:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_enhanced_critical_alerts(lookback_days, high_risk_threshold, auto_block_threshold, time_window, internal_ips, exclude_internal, enable_auto_blocking)
                print '</div>'
            
            # ::STEP:: Generate Executive Summary
            if show_summary:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_enhanced_summary_statistics(lookback_days, max_results)
                print '</div>'
            
            # ::STEP:: Generate Recent Activity Timeline
            if show_recent_activity:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_recent_activity_timeline(recent_events_limit, internal_ips, exclude_internal)
                print '</div>'
            
            # ::STEP:: Generate Advanced Threat Analysis
            if show_threat_analysis:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_advanced_threat_analysis(lookback_days, high_risk_threshold, time_window, min_attempts, activity_filter, internal_ips, exclude_internal, enable_pattern_analysis)
                print '</div>'
            
            # ::STEP:: Generate Successful Login Comparison
            if show_successful_logins:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_enhanced_login_comparison(lookback_days)
                print '</div>'
            
            # ::STEP:: Generate Detailed Analysis
            if show_detailed:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_enhanced_detailed_analysis(lookback_days, min_attempts, max_results, activity_filter, internal_ips, exclude_internal, show_user_details)
                print '</div>'
            
            # ::STEP:: Generate IP Analysis
            if show_ip_analysis:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_enhanced_ip_analysis(lookback_days, min_attempts, max_results, internal_ips, exclude_internal, enable_geo_tracking)
                print '</div>'
            
            # ::STEP:: Generate User Behavior Analysis
            if show_user_analysis:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_user_behavior_analysis(lookback_days, min_attempts, max_results)
                print '</div>'
            
            # ::STEP:: Generate Compliance Report
            if show_compliance_report:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_compliance_audit_report(lookback_days, high_risk_threshold)
                print '</div>'
            
            # ::STEP:: Generate Remediation Recommendations
            if include_remediation:
                print '<div class="tpsecuritydash2025_section">'
                self.generate_remediation_recommendations(lookback_days, high_risk_threshold, auto_block_threshold)
                print '</div>'
            
            # ::STEP:: Generate Navigation and Export Options
            print """
                    </div>
                </div>
                <div class="tpsecuritydash2025_backbuttons">
                    <a href="javascript:history.back()" class="tpsecuritydash2025_backbtn">
                        <i class="fa fa-arrow-left"></i> Back to Configuration
                    </a>
                    <button onclick="window.print()" class="tpsecuritydash2025_backbtn tpsecuritydash2025_printbtn">
                        <i class="fa fa-print"></i> Print Report
                    </button>
                    <button onclick="exportCSV()" class="tpsecuritydash2025_backbtn tpsecuritydash2025_exportbtn">
                        <i class="fa fa-download"></i> Export CSV
                    </button>
                    {}
                </div>
            </div>
            """.format(
                '<button onclick="refreshRealtime()" class="tpsecuritydash2025_backbtn" style="background-color: #28a745;"><i class="fa fa-refresh"></i> Refresh Real-time</button>' if enable_realtime_monitoring else ''
            )
            
        except Exception as e:
            self.print_error("Enhanced Security Report Generation", e)
    
    def generate_enhanced_critical_alerts(self, lookback_days, high_risk_threshold, auto_block_threshold, time_window, internal_ips="", exclude_internal=False, enable_auto_blocking=False):
        """Generate enhanced critical security alerts with auto-blocking recommendations"""
        try:
            # ::START:: Build enhanced WHERE clause
            where_conditions = ["Activity LIKE '%failed%'"]
            where_conditions.append("ActivityDate >= DATEADD(DAY, -{}, GETDATE())".format(lookback_days))
            
            if exclude_internal and internal_ips.strip():
                ip_prefixes = [ip.strip() for ip in internal_ips.split(',') if ip.strip()]
                for prefix in ip_prefixes:
                    clean_prefix = prefix.rstrip('.')
                    where_conditions.append("ClientIp NOT LIKE '{}%'".format(clean_prefix))
            
            where_clause = " AND ".join(where_conditions)
            
            # ::START:: Enhanced Critical Threats Query with Users table integration
            sql_critical = """
                SELECT TOP 25
                    ClientIp,
                    COUNT(*) AS TotalAttempts,
                    COUNT(DISTINCT Activity) AS UniqueActivities,
                    COUNT(DISTINCT al.UserId) AS UsersTargeted,
                    COUNT(DISTINCT CAST(ActivityDate AS DATE)) AS DaysActive,
                    MIN(ActivityDate) AS FirstSeen,
                    MAX(ActivityDate) AS LastSeen,
                    DATEDIFF(MINUTE, MIN(ActivityDate), MAX(ActivityDate)) AS DurationMinutes,
                    -- Enhanced user account status information
                    SUM(CASE WHEN u.IsLockedOut = 1 THEN 1 ELSE 0 END) AS CurrentlyLockedUsers,
                    SUM(CASE WHEN u.LastLockedOutDate >= DATEADD(DAY, -{}, GETDATE()) THEN 1 ELSE 0 END) AS RecentlyLockedUsers,
                    SUM(CASE WHEN al.Activity LIKE '%passwordreset%' OR al.Activity LIKE '%ForgotPassword%' THEN 1 ELSE 0 END) AS PasswordResetAttempts,
                    SUM(CASE WHEN u.MustChangePassword = 1 THEN 1 ELSE 0 END) AS ForcedPasswordResets,
                    MAX(u.FailedPasswordAttemptCount) AS MaxFailedAttempts,
                    CASE 
                        WHEN COUNT(*) >= {} THEN 'EMERGENCY - Immediate Action Required'
                        WHEN COUNT(*) >= {} THEN 'CRITICAL - High Priority'
                        WHEN COUNT(*) >= {} THEN 'HIGH - Monitor Closely'
                        ELSE 'MEDIUM - Standard Monitoring'
                    END AS ThreatLevel,
                    CASE 
                        WHEN COUNT(*) >= {} AND DATEDIFF(MINUTE, MIN(ActivityDate), MAX(ActivityDate)) <= {} THEN 1
                        ELSE 0
                    END AS IsBruteForce
                FROM ActivityLog al
                LEFT JOIN Users u ON al.UserId = u.UserId
                WHERE {}
                GROUP BY ClientIp
                HAVING COUNT(*) >= {}
                ORDER BY TotalAttempts DESC, DaysActive ASC;
            """.format(
                lookback_days,  # For recently locked users check
                auto_block_threshold * 2,  # Emergency threshold
                auto_block_threshold,      # Critical threshold
                high_risk_threshold,       # High threshold
                high_risk_threshold,       # Brute force detection
                time_window,
                where_clause, 
                high_risk_threshold // 2   # Minimum for inclusion
            )
            
            critical_data = q.QuerySql(sql_critical)
            
            if critical_data and len(critical_data) > 0:
                # ::STEP:: Categorize threats by severity
                emergency_threats = []
                critical_threats = []
                high_threats = []
                brute_force_threats = []
                
                for threat in critical_data:
                    if threat.TotalAttempts >= auto_block_threshold * 2:
                        emergency_threats.append(threat)
                    elif threat.TotalAttempts >= auto_block_threshold:
                        critical_threats.append(threat)
                    elif threat.TotalAttempts >= high_risk_threshold:
                        high_threats.append(threat)
                    
                    if threat.IsBruteForce == 1:
                        brute_force_threats.append(threat)
                
                # ::STEP:: Display Emergency Alerts
                if emergency_threats:
                    print """
                    <div class="tpsecuritydash2025_criticalAlert tpsecuritydash2025_emergencyAlert">
                        <h2 style="color: #ffffff; margin-top: 0; font-size: 24px;">
                            <i class="fa fa-exclamation-triangle fa-2x"></i> EMERGENCY SECURITY ALERT
                        </h2>
                        <p style="font-size: 18px; font-weight: bold;">IMMEDIATE ACTION REQUIRED - ACTIVE CYBER ATTACK DETECTED</p>
                        <div style="background-color: rgba(255,255,255,0.2); padding: 15px; border-radius: 6px; margin: 15px 0;">
                            <h4 style="color: #ffffff; margin-top: 0;">Critical Threat Sources:</h4>
                            <ul style="font-size: 16px; margin: 0;">
                    """
                    
                    for threat in emergency_threats[:5]:
                        print """
                                <li><strong>IP {}:</strong> {} failed attempts across {} days targeting {} users</li>
                        """.format(threat.ClientIp, threat.TotalAttempts, threat.DaysActive, threat.UsersTargeted)
                    
                    if enable_auto_blocking:
                        print """
                            </ul>
                            <h4 style="color: #ffffff; margin: 20px 0 10px 0;">AUTOMATIC BLOCKING RECOMMENDATIONS:</h4>
                            <div style="background-color: rgba(255,255,255,0.3); padding: 10px; border-radius: 4px;">
                        """
                        for threat in emergency_threats[:3]:
                            print """
                                <div style="margin: 5px 0; padding: 8px; background-color: rgba(255,255,255,0.2); border-radius: 4px;">
                                    <strong>BLOCK IP {}:</strong> {} attempts in {} minutes
                                    <br><small>Command: iptables -A INPUT -s {} -j DROP</small>
                                </div>
                            """.format(threat.ClientIp, threat.TotalAttempts, threat.DurationMinutes, threat.ClientIp)
                        print """
                            </div>
                        """
                    
                    print """
                        </div>
                        <div style="background-color: rgba(255,255,255,0.2); padding: 15px; border-radius: 6px;">
                            <h4 style="color: #ffffff; margin-top: 0;">IMMEDIATE ACTIONS REQUIRED:</h4>
                            <ol style="font-size: 14px; margin: 0;">
                                <li><strong>BLOCK</strong> the identified IP addresses immediately at firewall level</li>
                                <li><strong>VERIFY</strong> all affected user accounts for compromise</li>
                                <li><strong>IMPLEMENT</strong> emergency rate limiting on authentication endpoints</li>
                                <li><strong>ACTIVATE</strong> incident response procedures</li>
                                <li><strong>NOTIFY</strong> security team and management immediately</li>
                                <li><strong>MONITOR</strong> for additional attack vectors and IP addresses</li>
                                <li><strong>DOCUMENT</strong> all actions taken for forensic analysis</li>
                            </ol>
                        </div>
                    </div>
                    """
                
                # ::STEP:: Display Critical and High Severity Alerts
                if critical_threats or high_threats:
                    print """
                    <div class="tpsecuritydash2025_criticalAlert">
                        <h3 style="color: #721c24; margin-top: 0; font-size: 20px;">
                            <i class="fa fa-shield"></i> CRITICAL SECURITY THREATS DETECTED
                        </h3>
                        <p><strong>Active security threats identified based on your data analysis:</strong></p>
                    """
                    
                    if critical_threats:
                        print """
                        <div style="margin: 15px 0;">
                            <h4 style="color: #721c24; margin-bottom: 10px;">Critical Priority Threats:</h4>
                            <ul style="margin: 0;">
                        """
                        for threat in critical_threats[:5]:
                            print """
                                <li><strong>IP {}:</strong> {} failed attempts over {} days, {} users targeted</li>
                            """.format(threat.ClientIp, threat.TotalAttempts, threat.DaysActive, threat.UsersTargeted)
                        print "</ul></div>"
                    
                    if high_threats:
                        print """
                        <div style="margin: 15px 0;">
                            <h4 style="color: #721c24; margin-bottom: 10px;">High Priority Threats:</h4>
                            <ul style="margin: 0;">
                        """
                        for threat in high_threats[:5]:
                            print """
                                <li><strong>IP {}:</strong> {} failed attempts over {} days</li>
                            """.format(threat.ClientIp, threat.TotalAttempts, threat.DaysActive)
                        print "</ul></div>"
                    
                    # ::STEP:: Display Brute Force Alerts
                    if brute_force_threats:
                        print """
                        <div style="background-color: #fff3cd; border: 2px solid #ffc107; padding: 15px; margin: 15px 0; border-radius: 6px;">
                            <h4 style="color: #856404; margin-top: 0;">
                                <i class="fa fa-bolt"></i> Brute Force Attacks Detected:
                            </h4>
                            <ul style="margin: 0; color: #856404;">
                        """
                        for threat in brute_force_threats[:3]:
                            print """
                                <li><strong>IP {}:</strong> {} rapid attempts in {} minutes</li>
                            """.format(threat.ClientIp, threat.TotalAttempts, threat.DurationMinutes)
                        print "</ul></div>"
                    
                    print """
                        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; margin-top: 15px;">
                            <h4 style="color: #495057; margin-top: 0;">PRIORITY ACTIONS REQUIRED:</h4>
                            <ol style="margin: 0; color: #495057;">
                                <li><strong>Review and block</strong> high-risk IP addresses immediately</li>
                                <li><strong>Check affected user accounts</strong> for signs of compromise</li>
                                <li><strong>Implement rate limiting</strong> on authentication endpoints</li>
                                <li><strong>Enable account lockout policies</strong> if not already active</li>
                                <li><strong>Consider IP-based geographic restrictions</strong> for known attack sources</li>
                                <li><strong>Monitor continuously</strong> for escalation or new attack vectors</li>
                                <li><strong>Update security policies</strong> based on attack patterns observed</li>
                            </ol>
                        </div>
                    </div>
                    """
            else:
                print """
                <div style="border: 2px solid #28a745; background-color: #d4edda; color: #155724; padding: 20px; margin: 20px 0; border-radius: 8px; font-family: Arial, sans-serif;">
                    <h3 style="color: #155724; margin-top: 0;">
                        <i class="fa fa-check-circle"></i> No Critical Security Threats Detected
                    </h3>
                    <p style="margin: 0;">No security incidents meeting critical alert thresholds were found in the analysis period. Continue regular monitoring.</p>
                </div>
                """
            
        except Exception as e:
            self.print_error("Enhanced Critical Alerts Generation", e)
            

    def generate_role_based_analysis(self, lookback_days, monitored_roles):
        """Generate analysis by role type"""
        try:
            print """
    <div style="border: 2px solid #6f42c1; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
        <div style="background: linear-gradient(135deg, #e2d9f3 0%, #d6c7f0 100%); border-bottom: 1px solid #c7b3ed; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #333333;">
            <i class="fa fa-users-cog"></i> Role-Based Security Risk Analysis (Last {} Days)
        </div>
        <div style="padding: 20px;">
            """.format(lookback_days)
            
            # ::STEP:: OPTIMIZED - Single query for all roles to avoid multiple table scans
            if monitored_roles:
                # Build role conditions for a single query
                role_conditions = []
                for role in monitored_roles:
                    role_conditions.append("ul.Roles LIKE '%{}%'".format(role))
                role_where = " OR ".join(role_conditions)
                
                sql_all_roles = """
                    SELECT 
                        -- Extract role name for grouping
                        CASE 
                            WHEN ul.Roles LIKE '%Admin%' THEN 'Admin'
                            WHEN ul.Roles LIKE '%Finance%' THEN 'Finance' 
                            WHEN ul.Roles LIKE '%FinanceAdmin%' THEN 'FinanceAdmin'
                            WHEN ul.Roles LIKE '%Developer%' THEN 'Developer'
                            WHEN ul.Roles LIKE '%ApiOnly%' THEN 'ApiOnly'
                            ELSE 'Other'
                        END AS RoleName,
                        
                        -- Count unique users with this role (total)
                        COUNT(DISTINCT ul.UserId) AS TotalUsersWithRole,
                        
                        -- Count users with failures in the period
                        COUNT(DISTINCT CASE WHEN al.ActivityDate IS NOT NULL THEN ul.UserId END) AS UsersWithFailures,
                        
                        -- Count total failures in the period
                        COUNT(al.ActivityDate) AS TotalFailures,
                        
                        -- Count unique IPs attacking users with this role
                        COUNT(DISTINCT al.ClientIp) AS UniqueIPs,
                        
                        -- Count unique days with incidents
                        COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) AS DaysWithIncidents,
                        
                        -- Current account status (not time-limited)
                        COUNT(DISTINCT CASE WHEN u.IsLockedOut = 1 THEN ul.UserId END) AS CurrentlyLockedUsers,
                        COUNT(DISTINCT CASE WHEN u.LastLockedOutDate >= DATEADD(DAY, -{}, GETDATE()) THEN ul.UserId END) AS RecentlyLockedUsers,
                        COUNT(DISTINCT CASE WHEN u.MustChangePassword = 1 THEN ul.UserId END) AS ForcedPasswordResets
                        
                    FROM UserList ul
                    LEFT JOIN ActivityLog al ON ul.UserId = al.UserId 
                        AND al.Activity LIKE '%failed%'
                        AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                    LEFT JOIN Users u ON ul.UserId = u.UserId
                    WHERE ({})
                    GROUP BY 
                        CASE 
                            WHEN ul.Roles LIKE '%Admin%' THEN 'Admin'
                            WHEN ul.Roles LIKE '%Finance%' THEN 'Finance'
                            WHEN ul.Roles LIKE '%FinanceAdmin%' THEN 'FinanceAdmin' 
                            WHEN ul.Roles LIKE '%Developer%' THEN 'Developer'
                            WHEN ul.Roles LIKE '%ApiOnly%' THEN 'ApiOnly'
                            ELSE 'Other'
                        END
                    ORDER BY TotalFailures DESC;
                """.format(lookback_days, lookback_days, role_where)
                
                try:
                    role_data = q.QuerySql(sql_all_roles)
                    if role_data and len(role_data) > 0:
                        for role_stats in role_data:
                            role_name = getattr(role_stats, 'RoleName', 'Unknown')
                            
                            # Skip if this role wasn't in our monitored list
                            if role_name not in monitored_roles and role_name != 'Other':
                                continue
                                
                            self.display_role_analysis_results(role_name, role_stats, lookback_days)
                    else:
                        print """
                        <div style="margin: 15px 0; padding: 15px; border: 1px solid #28a745; border-radius: 6px; background-color: #d4edda; color: #155724;">
                            <strong>No Security Incidents:</strong> No users with monitored roles experienced security incidents in the last {} days.
                        </div>
                        """.format(lookback_days)
                except Exception as role_error:
                    print """
                    <div style="margin: 15px 0; padding: 15px; border: 1px solid #dc3545; border-radius: 6px; background-color: #f8d7da;">
                        <h4 style="color: #721c24; margin-top: 0;">
                            Role Analysis Error
                        </h4>
                        <p style="color: #721c24; margin: 0;">
                            Error analyzing roles: {}
                        </p>
                    </div>
                    """.format(str(role_error)[:100])
            else:
                print """
                <div style="margin: 15px 0; padding: 15px; border: 1px solid #ffc107; border-radius: 6px; background-color: #fff3cd;">
                    <strong>No Roles Selected:</strong> No high-risk roles selected for monitoring.
                </div>
                """
        except Exception as e:
            self.print_error("Enhanced Critical Alerts Generation", e)
    
    def generate_recent_activity_timeline(self, recent_events_limit, internal_ips="", exclude_internal=False):
        """Generate recent security events timeline for active monitoring"""
        try:
            # ::START:: Build WHERE clause for recent events
            where_conditions = ["Activity LIKE '%failed%'"]
            
            if exclude_internal and internal_ips.strip():
                ip_prefixes = [ip.strip() for ip in internal_ips.split(',') if ip.strip()]
                for prefix in ip_prefixes:
                    clean_prefix = prefix.rstrip('.')
                    where_conditions.append("ClientIp NOT LIKE '{}%'".format(clean_prefix))
            
            where_clause = " AND ".join(where_conditions)
            
            # ::START:: Recent Events Query
            sql_recent = """
                SELECT TOP {}
                    ActivityDate,
                    Activity,
                    ClientIp,
                    ISNULL(u.Username, 'Unknown User') as Username,
                    u.PeopleId,
                    Machine,
                    CASE 
                        WHEN Activity LIKE '%ForgotPassword%' THEN 'Password Reset'
                        WHEN Activity LIKE '%invalid log in%' THEN 'Invalid Login'
                        WHEN Activity LIKE '%failed password%' THEN 'Failed Password'
                        WHEN Activity LIKE '%locked%' THEN 'Account Locked'
                        ELSE 'Other Failed Auth'
                    END AS EventType,
                    CASE 
                        WHEN Activity LIKE '%locked%' THEN 'Critical'
                        WHEN Activity LIKE '%ForgotPassword%' THEN 'High'
                        WHEN Activity LIKE '%failed password%' THEN 'Medium'
                        ELSE 'Low'
                    END AS Severity
                FROM ActivityLog al
                LEFT JOIN Users u ON al.UserId = u.UserId
                WHERE {}
                ORDER BY ActivityDate DESC;
            """.format(recent_events_limit, where_clause)
            
            recent_data = q.QuerySql(sql_recent)
            
            if recent_data and len(recent_data) > 0:
                print """
<div style="border: 2px solid #17a2b8; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
    <div style="background: linear-gradient(135deg, #d1ecf1 0%, #bee5eb 100%); border-bottom: 1px solid #abdde5; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #333333;">
        <i class="fa fa-clock-o"></i> Recent Security Events Timeline - Last {} Events
        <span style="float: right; font-size: 14px; font-weight: normal;">
            <i class="fa fa-refresh"></i> Live monitoring of failed authentication attempts
        </span>
    </div>
    <div style="padding: 20px; max-height: 600px; overflow-y: auto;">
                """.format(recent_events_limit)
                
                current_date = None
                event_count = 0
                
                for event in recent_data:
                    event_count += 1
                    event_date = self.format_dotnet_datetime(event.ActivityDate)[:10]  # Get date part
                    event_time = self.format_dotnet_time(event.ActivityDate)
                    
                    # Add date separator
                    if current_date != event_date:
                        if current_date is not None:
                            print "</div>"  # Close previous date group
                        
                        current_date = event_date
                        print """
                        <div style="margin: 20px 0 15px 0; padding: 8px 12px; background-color: #e9ecef; border-left: 4px solid #6c757d; font-weight: bold; color: #495057;">
                            <i class="fa fa-calendar"></i> {}
                        </div>
                        <div style="margin-left: 20px;">
                        """.format(event_date)
                    
                    # Determine severity styling
                    severity = getattr(event, 'Severity', 'Low')
                    event_type = getattr(event, 'EventType', 'Unknown')
                    
                    if severity == 'Critical':
                        severity_color = 'background-color: #dc3545; color: #ffffff;'
                        timeline_color = '#dc3545'
                    elif severity == 'High':
                        severity_color = 'background-color: #ffc107; color: #000000;'
                        timeline_color = '#ffc107'
                    elif severity == 'Medium':
                        severity_color = 'background-color: #fd7e14; color: #ffffff;'
                        timeline_color = '#fd7e14'
                    else:
                        severity_color = 'background-color: #6c757d; color: #ffffff;'
                        timeline_color = '#6c757d'
                    
                    # Extract username from activity if available
                    username = getattr(event, 'Username', 'Unknown')
                    people_id = getattr(event, 'PeopleId', 0)
                    client_ip = getattr(event, 'ClientIp', 'Unknown')
                    machine = getattr(event, 'Machine', 'Unknown')
                    
                    print """
                    <div style="margin: 8px 0; padding: 12px; border: 1px solid #dee2e6; border-left: 4px solid {}; border-radius: 4px; background-color: #f8f9fa;">
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <span style="font-weight: bold; color: #333; margin-right: 10px;">{}</span>
                            <span style="display: inline-block; padding: 2px 6px; font-size: 11px; font-weight: bold; border-radius: 3px; {}">{}</span>
                            <span style="margin-left: auto; font-size: 12px; color: #6c757d;">#{}</span>
                        </div>
                        <div style="font-size: 13px; color: #495057; margin-bottom: 6px;">
                            <strong>Event:</strong> {}
                        </div>
                        <div style="font-size: 12px; color: #6c757d;">
                            <strong>IP:</strong> 
                            <span class="tpsecuritydash2025_clickable" onclick="openIPLookup('{}')">
                                {}
                            </span>
                            {} | 
                            <strong>Machine:</strong> {} |
                            <strong>Details:</strong> 
                            <span class="tpsecuritydash2025_clickable" onclick="toggleDetails('event_{}')">
                                View Details <i class="fa fa-caret-down"></i>
                            </span>
                        </div>
                        <div id="event_{}" class="tpsecuritydash2025_details">
                            <strong>Full Activity Log:</strong><br>
                            <code style="background-color: #e9ecef; padding: 8px; border-radius: 4px; display: block; margin: 8px 0; word-wrap: break-word;">
                                {}
                            </code>
                            <strong>Event Details:</strong><br>
                            • <strong>Timestamp:</strong> {}<br>
                            • <strong>Source IP:</strong> {}<br>
                            • <strong>Machine Name:</strong> {}<br>
                            • <strong>Event Type:</strong> {}<br>
                            • <strong>Severity Level:</strong> {}
                        </div>
                    </div>
                    """.format(
                        timeline_color,
                        event_time,
                        severity_color, severity,
                        event_count,
                        event_type,
                        client_ip, client_ip,
                        '<a href="javascript:viewUser({})" style="color: #007bff; text-decoration: underline;">{}</a>'.format(people_id, username) if people_id and people_id != 0 else username,
                        machine,
                        event_count,
                        event_count,
                        getattr(event, 'Activity', 'Unknown activity'),
                        self.format_dotnet_datetime(event.ActivityDate),
                        client_ip,
                        machine,
                        event_type,
                        severity
                    )
                
                if current_date is not None:
                    print "</div>"  # Close last date group
                
                print """
        <div style="margin-top: 20px; padding: 15px; background-color: #e3f2fd; border-radius: 4px; border-left: 4px solid #2196f3;">
            <strong><i class="fa fa-info-circle"></i> Timeline Information:</strong> 
            <ul style="margin-top: 10px; margin-bottom: 0;">
                <li>Events are displayed in chronological order (most recent first)</li>
                <li>Click IP addresses to perform reputation lookups</li>
                <li>Click usernames to view user profiles (when available)</li>
                <li>Click "View Details" to see complete activity log entries</li>
                <li>Timeline auto-refreshes when real-time monitoring is enabled</li>
            </ul>
        </div>
    </div>
</div>
                """
            else:
                print """
<div style="border: 2px solid #28a745; background-color: #d4edda; color: #155724; padding: 15px; margin: 20px 0; border-radius: 6px; font-family: Arial, sans-serif;">
    <strong>No Recent Activity:</strong> No recent security events found matching the specified criteria.
</div>
                """
                
        except Exception as e:
            self.print_error("Recent Activity Timeline Generation", e)
    
    def generate_enhanced_summary_statistics(self, lookback_days, max_results):
        """Generate enhanced executive summary with comprehensive KPIs"""
        try:
            # ::START:: Enhanced Summary Statistics Query with Users integration
            sql_summary = """
                DECLARE @StartDate DATETIME = DATEADD(DAY, -{}, GETDATE());
                
                SELECT 
                    'Enhanced Summary for last ' + CAST({} AS VARCHAR) + ' days' AS Period,
                    COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) AS DaysWithFailedLogins,
                    COUNT(DISTINCT al.ClientIp) AS UniqueIPs,
                    COUNT(*) AS TotalFailedAttempts,
                    COUNT(DISTINCT al.UserId) AS UniqueUsersTargeted,
                    COUNT(DISTINCT al.Activity) AS UniqueActivityTypes,
                    COUNT(DISTINCT al.Machine) AS UniqueMachines,
                    MIN(al.ActivityDate) AS EarliestAttempt,
                    MAX(al.ActivityDate) AS LatestAttempt,
                    
                    -- Enhanced password and lockout metrics from Users table
                    SUM(CASE WHEN al.Activity LIKE '%ForgotPassword%' OR al.Activity LIKE '%passwordreset%' THEN 1 ELSE 0 END) AS PasswordResetAttempts,
                    SUM(CASE WHEN al.Activity LIKE '%locked%' THEN 1 ELSE 0 END) AS ActivityLogLockouts,
                    SUM(CASE WHEN al.Activity LIKE '%invalid%' THEN 1 ELSE 0 END) AS InvalidLogins,
                    SUM(CASE WHEN al.Activity LIKE '%mobile%' THEN 1 ELSE 0 END) AS MobileFailures,
                    
                    -- Current user account status from Users table
                    (SELECT COUNT(*) FROM Users WHERE IsLockedOut = 1) AS CurrentlyLockedAccounts,
                    (SELECT COUNT(*) FROM Users WHERE LastLockedOutDate >= @StartDate) AS RecentlyLockedAccounts,
                    (SELECT COUNT(*) FROM Users WHERE MustChangePassword = 1) AS ForcedPasswordResets,
                    (SELECT MAX(FailedPasswordAttemptCount) FROM Users) AS MaxFailedAttemptsPerUser
                    
                FROM ActivityLog al
                LEFT JOIN Users u ON al.UserId = u.UserId
                WHERE al.Activity LIKE '%failed%'
                    AND al.ActivityDate >= @StartDate;
            """.format(lookback_days, lookback_days)
            
            summary_data = q.QuerySql(sql_summary)
            
            if summary_data:
                # Handle both iterator and direct object access
                if hasattr(summary_data, '__iter__') and len(list(summary_data)) > 0:
                    summary_data = q.QuerySql(sql_summary)  # Re-query to reset iterator
                    summary = list(summary_data)[0]
                elif hasattr(summary_data, 'TotalFailedAttempts'):
                    summary = summary_data
                else:
                    summary = None
                
                if summary:
                    # Safely access properties with defaults and format them properly
                    total_attempts = getattr(summary, 'TotalFailedAttempts', 0)
                    unique_ips = getattr(summary, 'UniqueIPs', 0)
                    days_with_logins = getattr(summary, 'DaysWithFailedLogins', 0)
                    users_targeted = getattr(summary, 'UniqueUsersTargeted', 0)
                    unique_activities = getattr(summary, 'UniqueActivityTypes', 0)
                    unique_machines = getattr(summary, 'UniqueMachines', 0)
                    password_resets = getattr(summary, 'PasswordResetAttempts', 0)
                    account_lockouts = getattr(summary, 'ActivityLogLockouts', 0)
                    invalid_logins = getattr(summary, 'InvalidLogins', 0)
                    mobile_failures = getattr(summary, 'MobileFailures', 0)
                    currently_locked = getattr(summary, 'CurrentlyLockedAccounts', 0)
                    recently_locked = getattr(summary, 'RecentlyLockedAccounts', 0)
                    forced_resets = getattr(summary, 'ForcedPasswordResets', 0)
                    max_failed_per_user = getattr(summary, 'MaxFailedAttemptsPerUser', 0)
                    earliest_attempt = getattr(summary, 'EarliestAttempt', None)
                    latest_attempt = getattr(summary, 'LatestAttempt', None)
                    
                    # Calculate additional metrics with proper formatting
                    if days_with_logins > 0:
                        avg_attempts_per_day = float(total_attempts) / days_with_logins
                        avg_attempts_str = "{:.1f}".format(avg_attempts_per_day)
                    else:
                        avg_attempts_str = "0.0"
                    
                    if unique_ips > 0:
                        avg_attempts_per_ip = float(total_attempts) / unique_ips
                        avg_ip_str = "{:.1f}".format(avg_attempts_per_ip)
                    else:
                        avg_ip_str = "0.0"
                    
                    # Calculate coverage percentage
                    if lookback_days > 0:
                        coverage_percent = int((float(days_with_logins) / lookback_days) * 100)
                    else:
                        coverage_percent = 0
                    
                    # Format the max results display
                    if max_results > 0:
                        max_results_str = str(max_results)
                    else:
                        max_results_str = "Unlimited"
                
                    print """
    <div style="border: 2px solid #0d6efd; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
        <div style="background: linear-gradient(135deg, #cfe2ff 0%, #b6d7ff 100%); border-bottom: 1px solid #9ec5fe; padding: 12px 18px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 16px; color: #333333;">
            <i class="fa fa-dashboard"></i> Executive Security Dashboard - Enhanced KPIs (Last {} Days)
            <span style="float: right; font-size: 12px; font-weight: normal;">
                Max Results: {} | Generated: {}
            </span>
        </div>
        <div style="padding: 20px;">
            <!-- Primary KPIs Row -->
            <div class="row" style="margin-bottom: 20px;">
                <div class="col-md-3">
                    <div style="background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); color: white; padding: 18px; border-radius: 8px; text-align: center; box-shadow: 0 4px 8px rgba(220, 53, 69, 0.3); margin-bottom: 15px;">
                        <div style="font-size: 28px; font-weight: bold; margin-bottom: 6px;">{}</div>
                        <div style="font-size: 12px; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px;">Total Failed Attempts</div>
                        <div style="font-size: 10px; opacity: 0.8;">
                            <i class="fa fa-exclamation-triangle"></i> {} avg/day
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div style="background: linear-gradient(135deg, #fd7e14 0%, #e8590c 100%); color: white; padding: 18px; border-radius: 8px; text-align: center; box-shadow: 0 4px 8px rgba(253, 126, 20, 0.3); margin-bottom: 15px;">
                        <div style="font-size: 28px; font-weight: bold; margin-bottom: 6px;">{}</div>
                        <div style="font-size: 12px; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px;">Unique IP Sources</div>
                        <div style="font-size: 10px; opacity: 0.8;">
                            <i class="fa fa-globe"></i> {} avg attempts/IP
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div style="background: linear-gradient(135deg, #20c997 0%, #17a085 100%); color: white; padding: 18px; border-radius: 8px; text-align: center; box-shadow: 0 4px 8px rgba(32, 201, 151, 0.3); margin-bottom: 15px;">
                        <div style="font-size: 28px; font-weight: bold; margin-bottom: 6px;">{}</div>
                        <div style="font-size: 12px; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px;">Active Threat Days</div>
                        <div style="font-size: 10px; opacity: 0.8;">
                            <i class="fa fa-calendar"></i> {}% of period
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div style="background: linear-gradient(135deg, #6f42c1 0%, #5a359a 100%); color: white; padding: 18px; border-radius: 8px; text-align: center; box-shadow: 0 4px 8px rgba(111, 66, 193, 0.3); margin-bottom: 15px;">
                        <div style="font-size: 28px; font-weight: bold; margin-bottom: 6px;">{}</div>
                        <div style="font-size: 12px; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px;">Users Targeted</div>
                        <div style="font-size: 10px; opacity: 0.8;">
                            <i class="fa fa-users"></i> Accounts at Risk
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Secondary KPIs Row -->
            <div class="row" style="margin-bottom: 20px;">
                <div class="col-md-3">
                    <div style="background: linear-gradient(135deg, #e83e8c 0%, #d91a72 100%); color: white; padding: 15px; border-radius: 6px; text-align: center; box-shadow: 0 3px 6px rgba(232, 62, 140, 0.3);">
                        <div style="font-size: 22px; font-weight: bold; margin-bottom: 4px;">{}</div>
                        <div style="font-size: 11px; opacity: 0.9;">Password Resets</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div style="background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); color: white; padding: 15px; border-radius: 6px; text-align: center; box-shadow: 0 3px 6px rgba(220, 53, 69, 0.3);">
                        <div style="font-size: 22px; font-weight: bold; margin-bottom: 4px;">{}</div>
                        <div style="font-size: 11px; opacity: 0.9;">Account Lockouts</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div style="background: linear-gradient(135deg, #ffc107 0%, #e0a800 100%); color: #000; padding: 15px; border-radius: 6px; text-align: center; box-shadow: 0 3px 6px rgba(255, 193, 7, 0.3);">
                        <div style="font-size: 22px; font-weight: bold; margin-bottom: 4px;">{}</div>
                        <div style="font-size: 11px; opacity: 0.8;">Invalid Logins</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div style="background: linear-gradient(135deg, #17a2b8 0%, #138496 100%); color: white; padding: 15px; border-radius: 6px; text-align: center; box-shadow: 0 3px 6px rgba(23, 162, 184, 0.3);">
                        <div style="font-size: 22px; font-weight: bold; margin-bottom: 4px;">{}</div>
                        <div style="font-size: 11px; opacity: 0.9;">Mobile Failures</div>
                    </div>
                </div>
            </div>
            
            <!-- Technical Metrics Row -->
            <div class="row" style="margin-bottom: 20px;">
                <div class="col-md-4">
                    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; border-left: 4px solid #0d6efd;">
                        <div style="font-size: 12px; color: #6c757d; margin-bottom: 4px;">
                            <i class="fa fa-cogs"></i> Activity Types Detected
                        </div>
                        <div style="font-size: 20px; font-weight: bold; color: #212529;">
                            {} Different Types
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; border-left: 4px solid #28a745;">
                        <div style="font-size: 12px; color: #6c757d; margin-bottom: 4px;">
                            <i class="fa fa-desktop"></i> Unique Machines
                        </div>
                        <div style="font-size: 20px; font-weight: bold; color: #212529;">
                            {} Systems
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; border-left: 4px solid #6f42c1;">
                        <div style="font-size: 12px; color: #6c757d; margin-bottom: 4px;">
                            <i class="fa fa-clock-o"></i> Analysis Period
                        </div>
                        <div style="font-size: 20px; font-weight: bold; color: #212529;">
                            {} Days
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Account Status Row -->
            <div class="row" style="margin-bottom: 20px;">
                <div class="col-md-4">
                    <div style="background-color: #f8d7da; padding: 15px; border-radius: 6px; border-left: 4px solid #dc3545;">
                        <div style="font-size: 12px; color: #721c24; margin-bottom: 4px;">
                            <i class="fa fa-lock"></i> Currently Locked Accounts
                        </div>
                        <div style="font-size: 20px; font-weight: bold; color: #721c24;">
                            {} Accounts
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div style="background-color: #fff3cd; padding: 15px; border-radius: 6px; border-left: 4px solid #ffc107;">
                        <div style="font-size: 12px; color: #856404; margin-bottom: 4px;">
                            <i class="fa fa-exclamation-triangle"></i> Recently Locked
                        </div>
                        <div style="font-size: 20px; font-weight: bold; color: #856404;">
                            {} Accounts
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div style="background-color: #d1ecf1; padding: 15px; border-radius: 6px; border-left: 4px solid #17a2b8;">
                        <div style="font-size: 12px; color: #0c5460; margin-bottom: 4px;">
                            <i class="fa fa-key"></i> Forced Password Resets
                        </div>
                        <div style="font-size: 20px; font-weight: bold; color: #0c5460;">
                            {} Accounts
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Time Range Information -->
            <div class="row">
                <div class="col-md-6">
                    <div style="background-color: #e3f2fd; padding: 12px; border-radius: 6px; border-left: 3px solid #2196f3;">
                        <div style="font-size: 11px; color: #1565c0; margin-bottom: 3px;">
                            <i class="fa fa-play-circle"></i> First Security Event
                        </div>
                        <div style="font-size: 13px; font-weight: bold; color: #212529;">
                            {}
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div style="background-color: #ffebee; padding: 12px; border-radius: 6px; border-left: 3px solid #f44336;">
                        <div style="font-size: 11px; color: #c62828; margin-bottom: 3px;">
                            <i class="fa fa-stop-circle"></i> Latest Security Event
                        </div>
                        <div style="font-size: 13px; font-weight: bold; color: #212529;">
                            {}
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Executive Summary Box -->
            <div style="margin-top: 20px; padding: 18px; background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); border-radius: 6px; border-left: 4px solid #2196f3;">
                <div style="font-size: 13px; color: #1565c0; font-weight: bold; margin-bottom: 8px;">
                    <i class="fa fa-info-circle"></i> <strong>Executive Summary</strong>
                </div>
                <div style="font-size: 12px; color: #1565c0; line-height: 1.5;">
                    Over the past {} days, our security monitoring detected <strong>{} failed authentication attempts</strong> 
                    from <strong>{} unique IP sources</strong>, affecting <strong>{} user accounts</strong>. 
                    Security events occurred on <strong>{} out of {} days</strong> ({}% of the monitoring period), 
                    with an average of <strong>{} attempts per day</strong>. 
                    Key concerns include <strong>{} password reset attempts</strong> and <strong>{} account lockouts</strong>, 
                    indicating potential targeted attacks. Mobile authentication failures account for <strong>{} incidents</strong>.
                    Currently, <strong>{} accounts are locked</strong> and <strong>{} accounts require password resets</strong>.
                </div>
            </div>
        </div>
    </div>
                    """.format(
                        lookback_days,
                        max_results_str,
                        datetime.now().strftime('%Y-%m-%d %H:%M'),
                        total_attempts,  # No comma formatting for integers in format string
                        avg_attempts_str,
                        unique_ips,
                        avg_ip_str,
                        days_with_logins,
                        coverage_percent,
                        users_targeted,
                        password_resets,
                        account_lockouts,
                        invalid_logins,
                        mobile_failures,
                        unique_activities,
                        unique_machines,
                        lookback_days,
                        currently_locked,
                        recently_locked,
                        forced_resets,
                        self.format_dotnet_datetime(earliest_attempt),
                        self.format_dotnet_datetime(latest_attempt),
                        lookback_days,
                        total_attempts,
                        unique_ips,
                        users_targeted,
                        days_with_logins,
                        lookback_days,
                        coverage_percent,
                        avg_attempts_str,
                        password_resets,
                        account_lockouts,
                        mobile_failures,
                        currently_locked,
                        forced_resets
                    )
                else:
                    print """
                    <div class="alert alert-info">
                        <strong>No Data:</strong> No summary data could be retrieved.
                    </div>
                    """
            else:
                print """
                <div class="alert alert-info">
                    <strong>No Data:</strong> No failed login attempts found in the specified time period.
                </div>
                """
                
        except Exception as e:
            self.print_error("Enhanced Summary Statistics Generation", e)
    
    def generate_advanced_threat_analysis(self, lookback_days, high_risk_threshold, time_window, min_attempts, activity_filter, internal_ips="", exclude_internal=False, enable_pattern_analysis=False):
        """Generate advanced threat analysis with pattern detection"""
        try:
            # ::START:: Build dynamic WHERE clause
            where_conditions = ["Activity LIKE '{}'".format(activity_filter)]
            where_conditions.append("ActivityDate >= DATEADD(DAY, -{}, GETDATE())".format(lookback_days))
            
            if exclude_internal and internal_ips.strip():
                ip_prefixes = [ip.strip() for ip in internal_ips.split(',') if ip.strip()]
                for prefix in ip_prefixes:
                    clean_prefix = prefix.rstrip('.')
                    where_conditions.append("ClientIp NOT LIKE '{}%'".format(clean_prefix))
            
            where_clause = " AND ".join(where_conditions)
            
            # ::START:: Advanced Threat Analysis Query - Fixed for SQL Server compatibility
            sql_threat = """
                WITH ThreatAnalysis AS (
                    SELECT 
                        ClientIp,
                        CAST(ActivityDate AS DATE) AS AttackDate,
                        COUNT(*) AS DailyAttempts,
                        COUNT(DISTINCT UserId) AS UsersTargeted,
                        COUNT(DISTINCT Activity) AS AttackTypes,
                        MIN(ActivityDate) AS FirstAttempt,
                        MAX(ActivityDate) AS LastAttempt,
                        DATEDIFF(SECOND, MIN(ActivityDate), MAX(ActivityDate)) AS DurationSeconds
                    FROM ActivityLog
                    WHERE {}
                    GROUP BY ClientIp, CAST(ActivityDate AS DATE)
                    HAVING COUNT(*) >= {}
                ),
                ThreatSummary AS (
                    SELECT 
                        ClientIp,
                        COUNT(*) AS DaysActive,
                        SUM(DailyAttempts) AS TotalAttempts,
                        AVG(DailyAttempts) AS AvgAttemptsPerDay,
                        MAX(DailyAttempts) AS PeakDailyAttempts,
                        SUM(UsersTargeted) AS TotalUsersTargeted,
                        MAX(UsersTargeted) AS MaxUsersPerDay,
                        COUNT(DISTINCT AttackTypes) AS UniqueAttackTypes,
                        MIN(AttackDate) AS FirstSeenDate,
                        MAX(AttackDate) AS LastSeenDate,
                        CASE 
                            WHEN MAX(DailyAttempts) >= {} AND MIN(DurationSeconds) <= {} THEN 'CRITICAL - Coordinated Attack'
                            WHEN MAX(DailyAttempts) >= {} THEN 'HIGH - Intensive Targeting'
                            WHEN COUNT(*) >= 3 THEN 'MEDIUM - Persistent Threat'
                            ELSE 'LOW - Standard Monitoring'
                        END AS ThreatClassification
                    FROM ThreatAnalysis
                    GROUP BY ClientIp
                )
                SELECT TOP 50
                    ts.ClientIp,
                    ts.DaysActive,
                    ts.TotalAttempts,
                    ts.AvgAttemptsPerDay,
                    ts.PeakDailyAttempts,
                    ts.TotalUsersTargeted,
                    ts.UniqueAttackTypes,
                    ts.FirstSeenDate,
                    ts.LastSeenDate,
                    ts.ThreatClassification,
                    DATEDIFF(DAY, ts.FirstSeenDate, ts.LastSeenDate) + 1 AS AttackSpanDays
                FROM ThreatSummary ts
                ORDER BY ts.TotalAttempts DESC, ts.DaysActive DESC;
            """.format(where_clause, min_attempts, high_risk_threshold * 2, time_window * 60, high_risk_threshold)
            
            threat_data = q.QuerySql(sql_threat)
            
            if threat_data and len(threat_data) > 0:
                print """
<div style="border: 2px solid #e83e8c; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
    <div style="background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%); border-bottom: 1px solid #f1aeb5; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 18px; color: #333333;">
        <i class="fa fa-crosshairs"></i> Advanced Threat Analysis - Pattern Detection & Intelligence
        {}
    </div>
    <div style="padding: 20px; overflow-x: auto;">
        <table style="width: 100%; border-collapse: collapse; margin: 0; font-size: 12px; font-family: Arial, sans-serif;">
            <thead>
                <tr>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">IP Address</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Threat Level</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Total Attempts</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Days Active</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Peak Daily</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Users Targeted</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Attack Types</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Time Span</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Pattern Analysis</th>
                </tr>
            </thead>
            <tbody>
                """.format(
                    '<span style="float: right; font-size: 14px; font-weight: normal; color: #28a745;"><i class="fa fa-cogs"></i> Advanced Pattern Detection Enabled</span>' if enable_pattern_analysis else ''
                )
                
                row_counter = 0
                for threat in threat_data:
                    row_counter += 1
                    
                    # Determine threat level styling
                    threat_classification = getattr(threat, 'ThreatClassification', 'LOW')
                    if 'CRITICAL' in threat_classification:
                        row_bg = 'background-color: #f8d7da;'
                        threat_color = 'background-color: #dc3545; color: #ffffff;'
                    elif 'HIGH' in threat_classification:
                        row_bg = 'background-color: #fff3cd;'
                        threat_color = 'background-color: #ffc107; color: #000000;'
                    elif 'MEDIUM' in threat_classification:
                        row_bg = ''
                        threat_color = 'background-color: #fd7e14; color: #ffffff;'
                    else:
                        row_bg = ''
                        threat_color = 'background-color: #6c757d; color: #ffffff;'
                    
                    # Calculate additional metrics
                    avg_attempts = getattr(threat, 'AvgAttemptsPerDay', 0)
                    peak_attempts = getattr(threat, 'PeakDailyAttempts', 0)
                    days_active = getattr(threat, 'DaysActive', 0)
                    attack_span = getattr(threat, 'AttackSpanDays', 0)
                    
                    # Format average attempts properly for display
                    if days_active > 0:
                        avg_attempts_display = "{:.1f}".format(float(avg_attempts))
                        avg_attempts_int = int(float(avg_attempts))
                    else:
                        avg_attempts_display = str(threat.TotalAttempts)
                        avg_attempts_int = threat.TotalAttempts
                    
                    # Pattern analysis
                    if enable_pattern_analysis:
                        pattern_analysis = self.analyze_attack_pattern(
                            threat.TotalAttempts, 
                            threat.DaysActive, 
                            threat.PeakDailyAttempts,
                            threat.UniqueAttackTypes,
                            attack_span
                        )
                    else:
                        pattern_analysis = "Basic Analysis"
                    
                    print """
                    <tr style="{}">
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tpsecuritydash2025_clickable" 
                                  onclick="openIPLookup('{}')" 
                                  style="font-family: monospace; background-color: #f8f9fa; padding: 3px 6px; border-radius: 3px; font-size: 11px; font-weight: bold;">
                                {}
                            </span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span style="display: inline-block; padding: 4px 8px; font-size: 10px; font-weight: bold; border-radius: 4px; white-space: nowrap; {}">{}</span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tpsecuritydash2025_clickable" onclick="toggleDetails('threat_total_{}')">
                                <strong>{:,}</strong> <i class="fa fa-caret-down"></i>
                            </span>
                            <div id="threat_total_{}" class="tpsecuritydash2025_details">
                                <strong>Attack Volume Analysis:</strong><br>
                                • <strong>Total Attempts:</strong> {:,}<br>
                                • <strong>Average per Day:</strong> {}<br>
                                • <strong>Peak Daily Volume:</strong> {:,}<br>
                                • <strong>Attack Intensity:</strong> {} attempts/day average
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{:,}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">
                            <span class="tpsecuritydash2025_clickable" onclick="toggleDetails('threat_users_{}')">
                                {:,} <i class="fa fa-caret-down"></i>
                            </span>
                            <div id="threat_users_{}" class="tpsecuritydash2025_details">
                                <strong>Target Analysis for {}:</strong><br>
                                • <strong>Total Users Targeted:</strong> {:,}<br>
                                • <strong>Attack Scope:</strong> {} different user accounts<br>
                                • <strong>Targeting Pattern:</strong> {} users per attempt average
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; font-size: 11px;">
                            {} to {}
                            <br><small>({} days span)</small>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tpsecuritydash2025_clickable" onclick="toggleDetails('pattern_{}')">
                                Pattern <i class="fa fa-caret-down"></i>
                            </span>
                            <div id="pattern_{}" class="tpsecuritydash2025_details">
                                <strong>Attack Pattern Analysis:</strong><br>
                                {}
                            </div>
                        </td>
                    </tr>
                    """.format(
                        row_bg,
                        threat.ClientIp,
                        threat.ClientIp,
                        threat_color, threat_classification,
                        row_counter,
                        threat.TotalAttempts,
                        row_counter,
                        threat.TotalAttempts,
                        avg_attempts_display,  # Changed from avg_attempts with format specifier
                        peak_attempts,
                        avg_attempts_int,  # Changed from int(avg_attempts)
                        days_active,
                        peak_attempts,
                        row_counter,
                        threat.TotalUsersTargeted,
                        row_counter,
                        threat.ClientIp,
                        threat.TotalUsersTargeted,
                        threat.TotalUsersTargeted,
                        "{:.2f}".format(float(threat.TotalUsersTargeted) / threat.TotalAttempts if threat.TotalAttempts > 0 else 0),
                        threat.UniqueAttackTypes,
                        self.format_dotnet_datetime_short(threat.FirstSeenDate),
                        self.format_dotnet_datetime_short(threat.LastSeenDate),
                        attack_span,
                        row_counter,
                        row_counter,
                        pattern_analysis
                    )
                
                print """
            </tbody>
        </table>
        
        <div style="margin-top: 20px; padding: 18px; background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-radius: 6px; border-left: 4px solid #ffc107;">
            <strong><i class="fa fa-lightbulb-o"></i> Advanced Threat Intelligence:</strong> 
            <ul style="margin-top: 10px; margin-bottom: 0; font-size: 13px;">
                <li><strong>Threat Classification:</strong> Automated classification based on volume, persistence, and attack patterns</li>
                <li><strong>Pattern Analysis:</strong> Identifies coordinated attacks, distributed threats, and persistent adversaries</li>
                <li><strong>Intelligence Integration:</strong> Click IP addresses for reputation lookup and threat intelligence</li>
                <li><strong>Temporal Analysis:</strong> Tracks attack evolution and identifies escalation patterns</li>
                <li><strong>Target Analysis:</strong> Correlates user targeting patterns with attack methodologies</li>
            </ul>
        </div>
    </div>
</div>
                """
            else:
                print """
<div style="border: 2px solid #28a745; background-color: #d4edda; color: #155724; padding: 15px; margin: 20px 0; border-radius: 6px; font-family: Arial, sans-serif;">
    <strong>No Advanced Threats:</strong> No sophisticated threat patterns detected in the analysis period.
</div>
                """
                
        except Exception as e:
            self.print_error("Advanced Threat Analysis Generation", e)
    
    def analyze_attack_pattern(self, total_attempts, days_active, peak_daily, attack_types, span_days):
        """Analyze attack patterns for threat intelligence"""
        try:
            analysis_parts = []
            
            # Volume analysis
            if total_attempts >= 100:
                analysis_parts.append("• <strong>High Volume Attack:</strong> {} total attempts".format(total_attempts))
            
            # Persistence analysis
            if days_active >= 7:
                analysis_parts.append("• <strong>Persistent Threat:</strong> Active for {} days".format(days_active))
            elif span_days > days_active * 2:
                analysis_parts.append("• <strong>Intermittent Campaign:</strong> Sporadic activity over {} days".format(span_days))
            
            # Intensity analysis
            if peak_daily >= 50:
                analysis_parts.append("• <strong>Burst Attack Pattern:</strong> Peak of {} attempts in single day".format(peak_daily))
            
            # Sophistication analysis
            if attack_types >= 5:
                analysis_parts.append("• <strong>Multi-Vector Attack:</strong> {} different attack methods".format(attack_types))
            elif attack_types >= 3:
                analysis_parts.append("• <strong>Coordinated Attack:</strong> Multiple attack vectors".format(attack_types))
            
            # Calculate attack rate
            if days_active > 0:
                avg_rate = float(total_attempts) / days_active
                if avg_rate >= 20:
                    analysis_parts.append("• <strong>Aggressive Rate:</strong> {:.1f} attempts per day average".format(avg_rate))
            
            # Pattern classification
            if total_attempts >= 200 and days_active >= 5:
                analysis_parts.append("• <strong>Classification:</strong> Advanced Persistent Threat (APT) characteristics")
            elif peak_daily >= 50 and span_days <= 2:
                analysis_parts.append("• <strong>Classification:</strong> Rapid brute force campaign")
            elif days_active >= 14:
                analysis_parts.append("• <strong>Classification:</strong> Long-term reconnaissance campaign")
            
            if not analysis_parts:
                analysis_parts.append("• <strong>Standard Attack:</strong> Basic attack pattern detected")
            
            return "<br>".join(analysis_parts)
            
        except Exception as e:
            return "• <strong>Analysis Error:</strong> Could not determine attack pattern"
    
    def generate_enhanced_detailed_analysis(self, lookback_days, min_attempts, max_results, activity_filter="%failed%", internal_ips="", exclude_internal=False, show_user_details=False):
        """Generate enhanced detailed analysis with improved performance and features"""
        try:
            # ::START:: Build dynamic WHERE clause with enhanced filtering
            where_conditions = ["Activity LIKE '{}'".format(activity_filter)]
            where_conditions.append("ActivityDate >= DATEADD(DAY, -{}, GETDATE())".format(lookback_days))
            
            if exclude_internal and internal_ips.strip():
                ip_prefixes = [ip.strip() for ip in internal_ips.split(',') if ip.strip()]
                for prefix in ip_prefixes:
                    clean_prefix = prefix.rstrip('.')
                    where_conditions.append("ClientIp NOT LIKE '{}%'".format(clean_prefix))
            
            where_clause = " AND ".join(where_conditions)
            
            # ::START:: Enhanced Detailed Analysis Query with performance optimization
            limit_clause = "TOP {}".format(max_results) if max_results > 0 else ""
            
            sql_detailed = """
                SELECT {}
                    CAST(ActivityDate AS DATE) AS LoginDay,
                    Activity,
                    ClientIp,
                    COUNT(*) AS AttemptCount,
                    MIN(ActivityDate) AS FirstAttempt,
                    MAX(ActivityDate) AS LastAttempt,
                    COUNT(DISTINCT UserId) AS UniqueUserIds,
                    COUNT(DISTINCT Machine) AS UniqueMachines,
                    CASE 
                        WHEN Activity LIKE '%ForgotPassword%' THEN 'Password Reset'
                        WHEN Activity LIKE '%invalid log in%' AND Activity LIKE '%incorrect password%' THEN 'Wrong Password'
                        WHEN Activity LIKE '%invalid log in%' AND Activity LIKE '%no user found%' THEN 'Invalid Username'
                        WHEN Activity LIKE '%locked%' THEN 'Account Lockout'
                        WHEN Activity LIKE '%mobile%' THEN 'Mobile Authentication'
                        ELSE 'Other Failed Auth'
                    END AS EventCategory,
                    CASE 
                        WHEN COUNT(*) >= 25 THEN 'Critical'
                        WHEN COUNT(*) >= 15 THEN 'High'
                        WHEN COUNT(*) >= 10 THEN 'Medium'
                        ELSE 'Low'
                    END AS RiskLevel
                FROM ActivityLog
                WHERE {}
                GROUP BY 
                    CAST(ActivityDate AS DATE),
                    Activity,
                    ClientIp
                HAVING COUNT(*) >= {}
                ORDER BY 
                    LoginDay DESC,
                    AttemptCount DESC;
            """.format(limit_clause, where_clause, min_attempts)
            
            detailed_data = q.QuerySql(sql_detailed)
            
            if detailed_data and len(detailed_data) > 0:
                print """
<div style="border: 2px solid #17a2b8; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
    <div style="background: linear-gradient(135deg, #d1ecf1 0%, #bee5eb 100%); border-bottom: 1px solid #abdde5; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 16px; color: #333333;">
        <i class="fa fa-list-alt"></i> Enhanced Detailed Security Events Analysis
        <span style="float: right; font-size: 12px; font-weight: normal;">
            Showing up to {} results | Click elements for detailed information
        </span>
    </div>
    <div style="padding: 20px; overflow-x: auto; max-height: 800px; overflow-y: auto;">
        <table style="width: 100%; border-collapse: collapse; margin: 0; font-size: 12px; font-family: Arial, sans-serif;">
            <thead style="position: sticky; top: 0; background-color: #f8f9fa; z-index: 10;">
                <tr>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Date</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Event Category</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">IP Address</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Attempts</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Risk Level</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Users</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Time Span</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Actions</th>
                </tr>
            </thead>
            <tbody>
                """.format(max_results if max_results > 0 else "All")
                
                row_counter = 0
                for detail in detailed_data:
                    row_counter += 1
                    time_span = self.calculate_dotnet_time_span(detail.FirstAttempt, detail.LastAttempt)
                    
                    # Enhanced risk level styling
                    risk_level = getattr(detail, 'RiskLevel', 'Low')
                    event_category = getattr(detail, 'EventCategory', 'Unknown')
                    
                    if risk_level == 'Critical':
                        row_bg = 'background-color: #f8d7da;'
                        risk_color = 'background-color: #dc3545; color: #ffffff;'
                    elif risk_level == 'High':
                        row_bg = 'background-color: #fff3cd;'
                        risk_color = 'background-color: #ffc107; color: #000000;'
                    elif risk_level == 'Medium':
                        row_bg = ''
                        risk_color = 'background-color: #fd7e14; color: #ffffff;'
                    else:
                        row_bg = ''
                        risk_color = 'background-color: #6c757d; color: #ffffff;'
                    
                    # Category-specific styling
                    if event_category == 'Account Lockout':
                        category_color = 'background-color: #dc3545; color: #ffffff;'
                    elif event_category == 'Password Reset':
                        category_color = 'background-color: #ffc107; color: #000000;'
                    elif event_category == 'Wrong Password':
                        category_color = 'background-color: #fd7e14; color: #ffffff;'
                    elif event_category == 'Invalid Username':
                        category_color = 'background-color: #6f42c1; color: #ffffff;'
                    else:
                        category_color = 'background-color: #6c757d; color: #ffffff;'
                    
                    print """
                    <tr style="{}">
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; font-weight: bold;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span style="display: inline-block; padding: 3px 8px; font-size: 10px; font-weight: bold; border-radius: 4px; {}">{}</span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tpsecuritydash2025_clickable" 
                                  onclick="openIPLookup('{}')" 
                                  style="font-family: monospace; background-color: #f8f9fa; padding: 3px 6px; border-radius: 3px; font-weight: bold;">
                                {}
                            </span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tpsecuritydash2025_clickable" 
                                  onclick="toggleDetails('enhanced_attempts_{}')" 
                                  style="display: inline-block; padding: 4px 10px; font-size: 12px; font-weight: bold; border-radius: 4px; min-width: 30px; text-align: center; background-color: #007bff; color: #ffffff;">
                                {} <i class="fa fa-info-circle"></i>
                            </span>
                            <div id="enhanced_attempts_{}" class="tpsecuritydash2025_details">
                                <strong>Enhanced Attack Analysis:</strong><br>
                                • <strong>Date:</strong> {}<br>
                                • <strong>Event Category:</strong> {}<br>
                                • <strong>Source IP:</strong> {}<br>
                                • <strong>Attack Window:</strong> {} to {}<br>
                                • <strong>Duration:</strong> {}<br>
                                • <strong>Intensity:</strong> {} attempts over this period<br>
                                • <strong>Target Analysis:</strong> {} users targeted from {} machines
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span style="display: inline-block; padding: 3px 8px; font-size: 10px; font-weight: bold; border-radius: 4px; {}">{}</span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">
                            <span class="tpsecuritydash2025_clickable" onclick="toggleDetails('enhanced_users_{}')">
                                {} <i class="fa fa-users"></i>
                            </span>
                            <div id="enhanced_users_{}" class="tpsecuritydash2025_details">
                                <strong>User Target Analysis:</strong><br>
                    """.format(
                        row_bg,
                        self.format_dotnet_datetime(detail.LoginDay),
                        category_color, event_category,
                        detail.ClientIp,
                        detail.ClientIp,
                        row_counter,
                        detail.AttemptCount,
                        row_counter,
                        self.format_dotnet_datetime(detail.LoginDay),
                        event_category,
                        detail.ClientIp,
                        self.format_dotnet_time(detail.FirstAttempt),
                        self.format_dotnet_time(detail.LastAttempt),
                        time_span,
                        detail.AttemptCount,
                        detail.UniqueUserIds,
                        detail.UniqueMachines,
                        risk_color, risk_level,
                        row_counter,
                        detail.UniqueUserIds,
                        row_counter
                    )
                    
                    # Enhanced user details with better error handling
                    if show_user_details:
                        try:
                            user_sql = """
                                SELECT DISTINCT 
                                    al.UserId, 
                                    ISNULL(u.Username, 'Unknown') as Username,
                                    u.PeopleId,
                                    COUNT(*) as UserAttempts
                                FROM ActivityLog al
                                LEFT JOIN Users u ON al.UserId = u.UserId
                                WHERE al.ClientIp = '{}'
                                    AND CAST(al.ActivityDate AS DATE) = '{}'
                                    AND al.Activity LIKE '%failed%'
                                    AND al.UserId IS NOT NULL
                                    AND al.UserId > 0
                                GROUP BY al.UserId, u.Username, u.PeopleId
                                ORDER BY UserAttempts DESC
                            """.format(detail.ClientIp, detail.LoginDay)
                            
                            user_data = q.QuerySql(user_sql)
                            user_count = 0
                            
                            if user_data:
                                for user in user_data:
                                    people_id = getattr(user, 'PeopleId', 0)
                                    username = getattr(user, 'Username', 'Unknown')
                                    user_attempts = getattr(user, 'UserAttempts', 0)
                                    
                                    if people_id and people_id != 0:
                                        print '• <a href="javascript:viewUser({})" style="color: #007bff; font-weight: bold; text-decoration: underline;">{}</a> ({} attempts)<br>'.format(people_id, username, user_attempts)
                                    else:
                                        print '• <span style="color: #dc3545; font-weight: bold;">{}</span> ({} attempts)<br>'.format(username, user_attempts)
                                    user_count += 1
                            
                            if user_count == 0:
                                print '• <em>No specific user data available - may indicate attacks using invalid usernames</em><br>'
                        except:
                            print '• <em>Error retrieving detailed user information</em><br>'
                    else:
                        print '• <em>Enable "Show User Details" in configuration for username information</em><br>'
                    
                    print """
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; font-size: 11px;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">
                            <div style="display: flex; gap: 4px; justify-content: center;">
                                <button onclick="openIPLookup('{}')" style="padding: 2px 6px; font-size: 10px; background-color: #17a2b8; color: white; border: none; border-radius: 3px; cursor: pointer;">
                                    <i class="fa fa-search"></i>
                                </button>
                                <button onclick="scrollToSection('ip-analysis-section')" style="padding: 2px 6px; font-size: 10px; background-color: #28a745; color: white; border: none; border-radius: 3px; cursor: pointer;">
                                    <i class="fa fa-link"></i>
                                </button>
                            </div>
                        </td>
                    </tr>
                    """.format(time_span, detail.ClientIp)
                
                print """
            </tbody>
        </table>
        
        <div style="margin-top: 20px; padding: 18px; background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); border-radius: 6px; border-left: 4px solid #2196f3;">
            <strong><i class="fa fa-info-circle"></i> Enhanced Interactive Analysis:</strong> 
            <ul style="margin-top: 10px; margin-bottom: 0; font-size: 13px;">
                <li><strong>Event Categories:</strong> Automated classification of security events by type and severity</li>
                <li><strong>Risk Assessment:</strong> Dynamic risk scoring based on attempt volume and patterns</li>
                <li><strong>Interactive Elements:</strong> Click IP addresses for reputation lookup, attempt counts for detailed analysis</li>
                <li><strong>User Intelligence:</strong> Click user counts to see specific targeted usernames (when enabled)</li>
                <li><strong>Cross-Reference:</strong> Use action buttons to cross-reference with other analysis sections</li>
                <li><strong>Performance:</strong> Results limited to {} entries for optimal loading performance</li>
            </ul>
        </div>
    </div>
</div>
                """.format(max_results if max_results > 0 else "unlimited")
            else:
                print """
<div style="border: 2px solid #28a745; background-color: #d4edda; color: #155724; padding: 15px; margin: 20px 0; border-radius: 6px; font-family: Arial, sans-serif;">
    <strong>No Detailed Events:</strong> No security incidents found matching the specified criteria.
</div>
                """
                
        except Exception as e:
            self.print_error("Enhanced Detailed Analysis Generation", e)
    
    def generate_enhanced_ip_analysis(self, lookback_days, min_attempts, max_results, internal_ips="", exclude_internal=False, enable_geo_tracking=False):
        """Generate enhanced IP analysis with geographic tracking and reputation intelligence"""
        try:
            # ::START:: Build enhanced WHERE clause
            where_conditions = ["Activity LIKE '%failed%'"]
            where_conditions.append("ActivityDate >= DATEADD(DAY, -{}, GETDATE())".format(lookback_days))
            
            if exclude_internal and internal_ips.strip():
                ip_prefixes = [ip.strip() for ip in internal_ips.split(',') if ip.strip()]
                for prefix in ip_prefixes:
                    clean_prefix = prefix.rstrip('.')
                    where_conditions.append("ClientIp NOT LIKE '{}%'".format(clean_prefix))
            
            where_clause = " AND ".join(where_conditions)
            limit_clause = "TOP {}".format(max_results) if max_results > 0 else ""
            
            # ::START:: Enhanced IP Analysis Query
            sql_ip = """
                SELECT {}
                    ClientIp,
                    COUNT(*) AS TotalAttempts,
                    COUNT(DISTINCT CAST(ActivityDate AS DATE)) AS DaysActive,
                    COUNT(DISTINCT Activity) AS UniqueActivities,
                    COUNT(DISTINCT UserId) AS UniqueUsers,
                    COUNT(DISTINCT Machine) AS UniqueMachines,
                    MIN(ActivityDate) AS FirstSeen,
                    MAX(ActivityDate) AS LastSeen,
                    AVG(CAST(COUNT(*) AS FLOAT)) OVER (PARTITION BY ClientIp) AS AvgAttemptsPerDay,
                    SUM(CASE WHEN Activity LIKE '%ForgotPassword%' THEN 1 ELSE 0 END) AS PasswordResets,
                    SUM(CASE WHEN Activity LIKE '%locked%' THEN 1 ELSE 0 END) AS AccountLockouts,
                    SUM(CASE WHEN Activity LIKE '%invalid%' THEN 1 ELSE 0 END) AS InvalidAttempts,
                    CASE 
                        WHEN COUNT(*) >= 100 THEN 'Critical Risk'
                        WHEN COUNT(*) >= 50 THEN 'High Risk'
                        WHEN COUNT(*) >= 25 THEN 'Medium Risk'
                        WHEN COUNT(*) >= 10 THEN 'Low Risk'
                        ELSE 'Minimal Risk'
                    END AS RiskLevel,
                    CASE 
                        WHEN COUNT(DISTINCT CAST(ActivityDate AS DATE)) >= 7 THEN 'Persistent'
                        WHEN COUNT(*) >= 50 AND COUNT(DISTINCT CAST(ActivityDate AS DATE)) <= 2 THEN 'Burst Attack'
                        WHEN COUNT(DISTINCT Activity) >= 5 THEN 'Multi-Vector'
                        ELSE 'Standard'
                    END AS AttackPattern
                FROM ActivityLog
                WHERE {}
                GROUP BY ClientIp
                HAVING COUNT(*) >= {}
                ORDER BY TotalAttempts DESC, DaysActive DESC;
            """.format(limit_clause, where_clause, min_attempts)
            
            ip_data = q.QuerySql(sql_ip)
            
            if ip_data and len(ip_data) > 0:
                print """
<div id="ip-analysis-section" style="border: 2px solid #6f42c1; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
    <div style="background: linear-gradient(135deg, #e2d9f3 0%, #d6c7f0 100%); border-bottom: 1px solid #c7b3ed; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 16px; color: #333333;">
        <i class="fa fa-globe"></i> Enhanced IP Address Intelligence & Analysis
        <span style="float: right; font-size: 12px; font-weight: normal;">
            {} | Geo-tracking: {} | Max Results: {}
        </span>
    </div>
    <div style="padding: 20px; overflow-x: auto; max-height: 700px; overflow-y: auto;">
        <table style="width: 100%; border-collapse: collapse; margin: 0; font-size: 12px; font-family: Arial, sans-serif;">
            <thead style="position: sticky; top: 0; background-color: #f8f9fa; z-index: 10;">
                <tr>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">IP Address</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Risk Level</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Total Attempts</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Activity Pattern</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Target Analysis</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Time Analysis</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Attack Types</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Intelligence</th>
                </tr>
            </thead>
            <tbody>
                """.format(
                    "Advanced IP Intelligence Analysis",
                    "Enabled" if enable_geo_tracking else "Disabled",
                    max_results if max_results > 0 else "All"
                )
                
                row_counter = 0
                for ip in ip_data:
                    row_counter += 1
                    
                    # Enhanced risk level styling
                    risk_level = getattr(ip, 'RiskLevel', 'Minimal Risk')
                    attack_pattern = getattr(ip, 'AttackPattern', 'Standard')
                    
                    if 'Critical' in risk_level:
                        row_bg = 'background-color: #f8d7da;'
                        risk_color = 'background-color: #dc3545; color: #ffffff;'
                    elif 'High' in risk_level:
                        row_bg = 'background-color: #fff3cd;'
                        risk_color = 'background-color: #ffc107; color: #000000;'
                    elif 'Medium' in risk_level:
                        row_bg = ''
                        risk_color = 'background-color: #fd7e14; color: #ffffff;'
                    elif 'Low' in risk_level:
                        row_bg = ''
                        risk_color = 'background-color: #6c757d; color: #ffffff;'
                    else:
                        row_bg = ''
                        risk_color = 'background-color: #28a745; color: #ffffff;'
                    
                    # Geographic analysis if enabled
                    geo_info = ""
                    if enable_geo_tracking:
                        geo_info = self.analyze_ip_geography(ip.ClientIp)
                    
                    # Calculate enhanced metrics
                    if ip.DaysActive > 0:
                        avg_attempts_per_day = float(ip.TotalAttempts) / ip.DaysActive
                        avg_attempts_str = "{:.1f}".format(avg_attempts_per_day)
                    else:
                        avg_attempts_str = str(ip.TotalAttempts)
                    
                    print """
                    <tr style="{}">
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <div style="display: flex; flex-direction: column; gap: 4px;">
                                <span class="tpsecuritydash2025_clickable" 
                                      onclick="openIPLookup('{}')" 
                                      style="font-family: monospace; background-color: #f8f9fa; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; border: 1px solid #dee2e6;">
                                    {} <i class="fa fa-external-link"></i>
                                </span>
                                {}
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span style="display: inline-block; padding: 4px 10px; font-size: 11px; font-weight: bold; border-radius: 4px; white-space: nowrap; {}">{}</span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tpsecuritydash2025_clickable" onclick="toggleDetails('enhanced_ip_attempts_{}')" style="display: inline-block; padding: 6px 12px; font-size: 13px; font-weight: bold; border-radius: 4px; min-width: 40px; text-align: center; background-color: #007bff; color: #ffffff;">
                                {:,} <i class="fa fa-chart-line"></i>
                            </span>
                            <div id="enhanced_ip_attempts_{}" class="tpsecuritydash2025_details">
                                <strong>Enhanced Attack Metrics for {}:</strong><br>
                                • <strong>Total Attempts:</strong> {:,}<br>
                                • <strong>Daily Average:</strong> {} attempts/day<br>
                                • <strong>Peak Activity:</strong> Most active over {} days<br>
                                • <strong>Password Resets:</strong> {} attempts<br>
                                • <strong>Account Lockouts:</strong> {} incidents<br>
                                • <strong>Invalid Attempts:</strong> {} tries<br>
                                • <strong>Attack Efficiency:</strong> {:.1f}% of attempts targeted valid users
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tpsecuritydash2025_clickable" onclick="toggleDetails('pattern_analysis_{}')">
                                <strong>{}</strong> <i class="fa fa-caret-down"></i>
                            </span>
                            <div id="pattern_analysis_{}" class="tpsecuritydash2025_details">
                                <strong>Attack Pattern Analysis:</strong><br>
                                • <strong>Classification:</strong> {}<br>
                                • <strong>Activity Span:</strong> {} days active<br>
                                • <strong>Attack Vectors:</strong> {} different methods<br>
                                • <strong>Persistence Level:</strong> {} pattern<br>
                                • <strong>Target Diversity:</strong> {} users, {} machines<br>
                                • <strong>Sophistication:</strong> {} attack sophistication level
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tpsecuritydash2025_clickable" onclick="toggleDetails('target_analysis_{}')">
                                <i class="fa fa-users"></i> {} Users <i class="fa fa-caret-down"></i>
                            </span>
                            <div id="target_analysis_{}" class="tpsecuritydash2025_details">
                                <strong>Target Intelligence for {}:</strong><br>
                    """.format(
                        row_bg,
                        ip.ClientIp,
                        ip.ClientIp,
                        geo_info,
                        risk_color, risk_level,
                        row_counter,
                        ip.TotalAttempts,
                        row_counter,
                        ip.ClientIp,
                        ip.TotalAttempts,
                        avg_attempts_str,
                        ip.DaysActive,
                        getattr(ip, 'PasswordResets', 0),
                        getattr(ip, 'AccountLockouts', 0),
                        getattr(ip, 'InvalidAttempts', 0),
                        (float(ip.UniqueUsers) / ip.TotalAttempts * 100) if ip.TotalAttempts > 0 else 0,
                        row_counter,
                        attack_pattern,
                        row_counter,
                        attack_pattern,
                        ip.DaysActive,
                        ip.UniqueActivities,
                        attack_pattern,
                        ip.UniqueUsers,
                        ip.UniqueMachines,
                        self.calculate_sophistication_level(ip.UniqueActivities, ip.TotalAttempts),
                        row_counter,
                        ip.UniqueUsers,
                        row_counter,
                        ip.ClientIp
                    )
                    
                    # Enhanced target analysis with user details
                    try:
                        user_sql = """
                            SELECT DISTINCT 
                                al.UserId, 
                                ISNULL(u.Username, 'Unknown') as Username,
                                u.PeopleId,
                                COUNT(*) as UserAttempts,
                                MAX(al.ActivityDate) as LastTargeted
                            FROM ActivityLog al
                            LEFT JOIN Users u ON al.UserId = u.UserId
                            WHERE al.ClientIp = '{}'
                                AND al.Activity LIKE '%failed%'
                                AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                                AND al.UserId IS NOT NULL
                                AND al.UserId > 0
                            GROUP BY al.UserId, u.Username, u.PeopleId
                            ORDER BY UserAttempts DESC
                        """.format(ip.ClientIp, lookback_days)
                        
                        user_data = q.QuerySql(user_sql)
                        valid_users = 0
                        
                        if user_data:
                            print '<strong>Valid User Accounts Targeted:</strong><br>'
                            for user in user_data:
                                people_id = getattr(user, 'PeopleId', 0)
                                username = getattr(user, 'Username', 'Unknown')
                                user_attempts = getattr(user, 'UserAttempts', 0)
                                last_targeted = getattr(user, 'LastTargeted', None)
                                
                                if people_id and people_id != 0:
                                    print '• <a href="javascript:viewUser({})" style="color: #dc3545; font-weight: bold; text-decoration: underline;">{}</a> ({} attempts, last: {})<br>'.format(
                                        people_id, username, user_attempts, self.format_dotnet_datetime_short(last_targeted)
                                    )
                                    valid_users += 1
                                else:
                                    print '• <span style="color: #6c757d;">{}</span> ({} attempts)<br>'.format(username, user_attempts)
                        
                        # Extract attempted usernames from activity logs
                        activity_sql = """
                            SELECT DISTINCT Activity, COUNT(*) as AttemptCount
                            FROM ActivityLog 
                            WHERE ClientIp = '{}'
                                AND Activity LIKE '%failed%'
                                AND ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                            GROUP BY Activity
                            ORDER BY AttemptCount DESC
                        """.format(ip.ClientIp, lookback_days)
                        
                        activity_data = q.QuerySql(activity_sql)
                        attempted_users = {}
                        
                        if activity_data:
                            for activity in activity_data:
                                activity_text = getattr(activity, 'Activity', '')
                                attempt_count = getattr(activity, 'AttemptCount', 0)
                                
                                extracted_username = self.extract_username_from_activity(activity_text)
                                if extracted_username:
                                    if extracted_username in attempted_users:
                                        attempted_users[extracted_username] += attempt_count
                                    else:
                                        attempted_users[extracted_username] = attempt_count
                            
                            if attempted_users:
                                print '<br><strong>All Attempted Usernames:</strong><br>'
                                sorted_attempts = sorted(attempted_users.items(), key=lambda x: x[1], reverse=True)
                                for username, total_attempts in sorted_attempts[:10]:
                                    print '• <span style="color: #fd7e14; font-weight: bold;">{}</span> ({} attempts)<br>'.format(username, total_attempts)
                        
                        print '<br><strong>Target Summary:</strong> {} valid users identified, {} total targeting attempts<br>'.format(
                            valid_users, ip.TotalAttempts
                        )
                        
                    except Exception as e:
                        print '• <em>Error analyzing target details: {}</em><br>'.format(str(e)[:50])
                    
                    print """
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; font-size: 11px;">
                            <div style="text-align: center;">
                                <strong>{}</strong> days<br>
                                <small>{} to {}</small>
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span class="tpsecuritydash2025_clickable" onclick="toggleDetails('attack_types_{}')">
                                {} Types <i class="fa fa-caret-down"></i>
                            </span>
                            <div id="attack_types_{}" class="tpsecuritydash2025_details">
                                <strong>Attack Vector Analysis:</strong><br>
                    """.format(
                        ip.DaysActive,
                        self.format_dotnet_datetime_short(ip.FirstSeen),
                        self.format_dotnet_datetime_short(ip.LastSeen),
                        row_counter,
                        ip.UniqueActivities,
                        row_counter
                    )
                    
                    # Get actual activity types for this IP
                    activities_sql = """
                        SELECT DISTINCT 
                            Activity,
                            COUNT(*) as ActivityCount
                        FROM ActivityLog 
                        WHERE ClientIp = '{}'
                            AND Activity LIKE '%failed%'
                            AND ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                        GROUP BY Activity
                        ORDER BY ActivityCount DESC
                    """.format(ip.ClientIp, lookback_days)
                    
                    try:
                        activities_data = q.QuerySql(activities_sql)
                        if activities_data:
                            for activity in activities_data:
                                activity_name = getattr(activity, 'Activity', 'Unknown')
                                activity_count = getattr(activity, 'ActivityCount', 0)
                                
                                # Categorize activity type
                                if 'forgotpassword' in activity_name.lower():
                                    activity_type = 'Password Reset'
                                    activity_color = '#ffc107'
                                elif 'invalid log in' in activity_name.lower():
                                    activity_type = 'Invalid Login'
                                    activity_color = '#dc3545'
                                elif 'locked' in activity_name.lower():
                                    activity_type = 'Account Lockout'
                                    activity_color = '#e83e8c'
                                else:
                                    activity_type = 'Other Auth Failure'
                                    activity_color = '#6c757d'
                                
                                print '• <span style="color: {}; font-weight: bold;">{}</span>: {} attempts<br>'.format(
                                    activity_color, activity_type, activity_count
                                )
                        else:
                            print '• <em>No specific activity data available</em><br>'
                    except:
                        print '• <em>Error retrieving activity breakdown</em><br>'
                    
                    print """
                            </div>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <div style="display: flex; flex-direction: column; gap: 4px;">
                                <button onclick="openIPLookup('{}')" style="padding: 4px 8px; font-size: 10px; background-color: #17a2b8; color: white; border: none; border-radius: 3px; cursor: pointer; width: 100%;">
                                    <i class="fa fa-search"></i> Reputation
                                </button>
                                <button onclick="toggleDetails('ip_intelligence_{}')" style="padding: 4px 8px; font-size: 10px; background-color: #6f42c1; color: white; border: none; border-radius: 3px; cursor: pointer; width: 100%;">
                                    <i class="fa fa-info"></i> Intel
                                </button>
                            </div>
                            <div id="ip_intelligence_{}" class="tpsecuritydash2025_details">
                                <strong>IP Intelligence Report:</strong><br>
                                • <strong>IP Address:</strong> {}<br>
                                • <strong>Risk Classification:</strong> {}<br>
                                • <strong>Attack Pattern:</strong> {}<br>
                                • <strong>Threat Level:</strong> Based on {} attempts over {} days<br>
                                • <strong>Persistence:</strong> {} day attack campaign<br>
                                • <strong>Sophistication:</strong> {} attack vectors detected<br>
                                {}
                                • <strong>Recommendation:</strong> {}
                            </div>
                        </td>
                    </tr>
                    """.format(
                        ip.ClientIp,
                        row_counter,
                        row_counter,
                        ip.ClientIp,
                        risk_level,
                        attack_pattern,
                        ip.TotalAttempts,
                        ip.DaysActive,
                        ip.DaysActive,
                        ip.UniqueActivities,
                        geo_info if geo_info else '',
                        self.get_ip_recommendation(risk_level, ip.TotalAttempts, attack_pattern)
                    )
                
                print """
            </tbody>
        </table>
        
        <div style="margin-top: 20px; padding: 20px; background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%); border-radius: 6px; border-left: 4px solid #9c27b0;">
            <strong><i class="fa fa-shield"></i> Enhanced IP Intelligence Summary:</strong> 
            <ul style="margin-top: 12px; margin-bottom: 0; font-size: 13px;">
                <li><strong>Risk Assessment:</strong> Automated risk scoring based on volume, persistence, and attack sophistication</li>
                <li><strong>Pattern Recognition:</strong> Identifies burst attacks, persistent threats, and multi-vector campaigns</li>
                <li><strong>Target Intelligence:</strong> Analyzes user targeting patterns and account enumeration attempts</li>
                <li><strong>Geographic Analysis:</strong> {} location tracking for enhanced threat attribution</li>
                <li><strong>Attack Vector Analysis:</strong> Categorizes and quantifies different attack methodologies</li>
                <li><strong>Reputation Integration:</strong> Direct links to IP reputation databases for threat verification</li>
                <li><strong>Actionable Intelligence:</strong> Provides specific recommendations for each identified threat</li>
            </ul>
        </div>
    </div>
</div>
                """.format("Advanced" if enable_geo_tracking else "Basic")
            else:
                print """
<div style="border: 2px solid #28a745; background-color: #d4edda; color: #155724; padding: 15px; margin: 20px 0; border-radius: 6px; font-family: Arial, sans-serif;">
    <strong>No IP Threats:</strong> No IP addresses found matching the specified threat criteria.
</div>
                """
                
        except Exception as e:
            self.print_error("Enhanced IP Analysis Generation", e)
    
    def generate_user_behavior_analysis(self, lookback_days, min_attempts, max_results):
        """Generate user behavior analysis for identifying compromised accounts"""
        try:
            sql_user_analysis = """
                SELECT {}
                    u.PeopleId,
                    u.Username,
                    u.IsLockedOut,
                    u.LastLockedOutDate,
                    u.FailedPasswordAttemptCount,
                    u.MustChangePassword,
                    u.LastLoginDate,
                    u.LastPasswordChangedDate,
                    COUNT(*) AS FailedAttempts,
                    COUNT(DISTINCT al.ClientIp) AS UniqueIPs,
                    COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) AS DaysWithFailures,
                    MIN(al.ActivityDate) AS FirstFailure,
                    MAX(al.ActivityDate) AS LastFailure,
                    COUNT(DISTINCT al.Activity) AS FailureTypes,
                    SUM(CASE WHEN al.Activity LIKE '%locked%' THEN 1 ELSE 0 END) AS ActivityLogLockouts,
                    SUM(CASE WHEN al.Activity LIKE '%passwordreset%' OR al.Activity LIKE '%ForgotPassword%' THEN 1 ELSE 0 END) AS PasswordResetRequests,
                    CASE 
                        WHEN u.IsLockedOut = 1 THEN 'CRITICAL - Account Locked'
                        WHEN COUNT(DISTINCT al.ClientIp) >= 10 THEN 'High Risk - Multiple IPs'
                        WHEN COUNT(*) >= 50 THEN 'High Risk - High Volume'
                        WHEN u.FailedPasswordAttemptCount >= 5 THEN 'Medium Risk - Failed Attempts'
                        WHEN u.MustChangePassword = 1 THEN 'Medium Risk - Password Reset Required'
                        ELSE 'Low Risk'
                    END AS RiskProfile
                FROM ActivityLog al
                JOIN Users u ON al.UserId = u.UserId
                WHERE al.Activity LIKE '%failed%'
                    AND al.ActivityDate >= DATEADD(DAY, -{}, GETDATE())
                    AND u.PeopleId IS NOT NULL
                GROUP BY u.PeopleId, u.Username, u.IsLockedOut, u.LastLockedOutDate, 
                         u.FailedPasswordAttemptCount, u.MustChangePassword, 
                         u.LastLoginDate, u.LastPasswordChangedDate
                HAVING COUNT(*) >= {}
                ORDER BY 
                    CASE WHEN u.IsLockedOut = 1 THEN 0 ELSE 1 END,  -- Locked accounts first
                    FailedAttempts DESC, 
                    UniqueIPs DESC;
            """.format(
                "TOP {}".format(max_results) if max_results > 0 else "",
                lookback_days,
                min_attempts
            )
            
            user_data = q.QuerySql(sql_user_analysis)
            
            if user_data and len(user_data) > 0:
                print """
<div style="border: 2px solid #e83e8c; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
    <div style="background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%); border-bottom: 1px solid #f1aeb5; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 16px; color: #333333;">
        <i class="fa fa-user-secret"></i> User Behavior Analysis - Potential Account Compromise Detection
    </div>
    <div style="padding: 20px; overflow-x: auto;">
        <table style="width: 100%; border-collapse: collapse; margin: 0; font-size: 12px; font-family: Arial, sans-serif;">
            <thead>
                <tr>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">User Account</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Risk Profile</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Failed Attempts</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Source IPs</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Time Analysis</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Security Events</th>
                    <th style="background-color: #f1f3f4; padding: 12px 8px; border: 1px solid #d1d3d4; font-weight: bold; text-align: left; font-size: 11px;">Actions</th>
                </tr>
            </thead>
            <tbody>
                """
                
                for user in user_data:
                    risk_profile = getattr(user, 'RiskProfile', 'Low Risk')
                    
                    if 'High Risk' in risk_profile:
                        row_bg = 'background-color: #f8d7da;'
                        risk_color = 'background-color: #dc3545; color: #ffffff;'
                    elif 'Medium Risk' in risk_profile:
                        row_bg = 'background-color: #fff3cd;'
                        risk_color = 'background-color: #ffc107; color: #000000;'
                    else:
                        row_bg = ''
                        risk_color = 'background-color: #28a745; color: #ffffff;'
                    
                    people_id = getattr(user, 'PeopleId', 0)
                    username = getattr(user, 'Username', 'Unknown')
                    
                    print """
                    <tr style="{}">
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <a href="javascript:viewUser({})" style="color: #007bff; font-weight: bold; text-decoration: underline;">
                                {} <i class="fa fa-external-link"></i>
                            </a>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle;">
                            <span style="display: inline-block; padding: 4px 10px; font-size: 10px; font-weight: bold; border-radius: 4px; {}">{}</span>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center; font-weight: bold;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">{}</td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; font-size: 11px;">
                            {} days<br>
                            <small>{} to {}</small>
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; font-size: 11px;">
                            {} types<br>
                            {} lockouts
                        </td>
                        <td style="padding: 10px 8px; border: 1px solid #d1d3d4; vertical-align: middle; text-align: center;">
                            <button onclick="viewUser({})" style="padding: 4px 8px; font-size: 10px; background-color: #007bff; color: white; border: none; border-radius: 3px; cursor: pointer; margin: 2px;">
                                <i class="fa fa-user"></i> Profile
                            </button>
                        </td>
                    </tr>
                    """.format(
                        row_bg,
                        people_id, username,
                        risk_color, risk_profile,
                        user.FailedAttempts,
                        user.UniqueIPs,
                        user.DaysWithFailures,
                        self.format_dotnet_datetime_short(user.FirstFailure),
                        self.format_dotnet_datetime_short(user.LastFailure),
                        user.FailureTypes,
                        user.AccountLockouts,
                        people_id
                    )
                
                print """
            </tbody>
        </table>
    </div>
</div>
                """
            else:
                print """
<div style="border: 2px solid #28a745; background-color: #d4edda; color: #155724; padding: 15px; margin: 20px 0; border-radius: 6px; font-family: Arial, sans-serif;">
    <strong>No User Behavior Concerns:</strong> No users showing suspicious authentication patterns.
</div>
                """
                
        except Exception as e:
            self.print_error("User Behavior Analysis Generation", e)
    
    def generate_compliance_audit_report(self, lookback_days, high_risk_threshold):
        """Generate compliance and audit report for security governance"""
        try:
            sql_compliance = """
                DECLARE @StartDate DATETIME = DATEADD(DAY, -{}, GETDATE());
                
                SELECT 
                    'Security Compliance Metrics' as ReportType,
                    COUNT(*) as TotalSecurityEvents,
                    COUNT(DISTINCT al.ClientIp) as UniqueThreats,
                    COUNT(DISTINCT al.UserId) as UsersAffected,
                    SUM(CASE WHEN al.Activity LIKE '%locked%' THEN 1 ELSE 0 END) as ActivityLogLockouts,
                    SUM(CASE WHEN al.Activity LIKE '%ForgotPassword%' OR al.Activity LIKE '%passwordreset%' THEN 1 ELSE 0 END) as PasswordResets,
                    COUNT(DISTINCT CAST(al.ActivityDate AS DATE)) as DaysWithIncidents,
                    
                    -- Account status metrics from Users table
                    (SELECT COUNT(*) FROM Users WHERE IsLockedOut = 1) AS CurrentlyLockedAccounts,
                    (SELECT COUNT(*) FROM Users WHERE LastLockedOutDate >= @StartDate) AS RecentlyLockedAccounts,
                    (SELECT COUNT(*) FROM Users WHERE MustChangePassword = 1) AS AccountsRequiringPasswordReset,
                    (SELECT COUNT(*) FROM Users WHERE FailedPasswordAttemptCount >= 3) AS AccountsWithMultipleFailures,
                    (SELECT COUNT(*) FROM Users WHERE LastLoginDate < DATEADD(DAY, -90, GETDATE())) AS InactiveAccounts
                    
                FROM ActivityLog al
                WHERE al.Activity LIKE '%failed%'
                    AND al.ActivityDate >= @StartDate;
            """.format(lookback_days)
            
            compliance_data = q.QuerySql(sql_compliance)
            
            if compliance_data:
                if hasattr(compliance_data, '__iter__') and len(list(compliance_data)) > 0:
                    compliance_data = q.QuerySql(sql_compliance)
                    compliance = list(compliance_data)[0]
                elif hasattr(compliance_data, 'TotalSecurityEvents'):
                    compliance = compliance_data
                else:
                    compliance = None
                
                if compliance:
                    # Safely access properties with defaults
                    total_events = getattr(compliance, 'TotalSecurityEvents', 0)
                    unique_threats = getattr(compliance, 'UniqueThreats', 0)
                    users_affected = getattr(compliance, 'UsersAffected', 0)
                    activity_lockouts = getattr(compliance, 'ActivityLogLockouts', 0)
                    password_resets = getattr(compliance, 'PasswordResets', 0)
                    days_with_incidents = getattr(compliance, 'DaysWithIncidents', 0)
                    currently_locked = getattr(compliance, 'CurrentlyLockedAccounts', 0)
                    recently_locked = getattr(compliance, 'RecentlyLockedAccounts', 0)
                    accounts_requiring_reset = getattr(compliance, 'AccountsRequiringPasswordReset', 0)
                    accounts_multiple_failures = getattr(compliance, 'AccountsWithMultipleFailures', 0)
                    inactive_accounts = getattr(compliance, 'InactiveAccounts', 0)
                    
                    # Calculate incident rate
                    if lookback_days > 0:
                        incident_rate = (float(days_with_incidents) / lookback_days * 100)
                        incident_rate_str = "{:.1f}".format(incident_rate)
                    else:
                        incident_rate_str = "0.0"
                    
                    # Determine security posture
                    if total_events > 0:
                        security_posture = "Active Monitoring"
                    else:
                        security_posture = "No Incidents"
                    
                    print """
    <div style="border: 2px solid #6f42c1; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
        <div style="background: linear-gradient(135deg, #e2d9f3 0%, #d6c7f0 100%); border-bottom: 1px solid #c7b3ed; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 16px; color: #333333;">
            <i class="fa fa-clipboard-check"></i> Security Compliance & Audit Report
        </div>
        <div style="padding: 20px;">
            <div class="row">
                <div class="col-md-6">
                    <h4 style="color: #6f42c1; margin-bottom: 15px;">Compliance Metrics</h4>
                    <ul style="list-style: none; padding: 0;">
                        <li style="margin: 10px 0; padding: 10px; background-color: #f8f9fa; border-left: 4px solid #6f42c1; border-radius: 4px;">
                            <strong>Total Security Events:</strong> {}
                        </li>
                        <li style="margin: 10px 0; padding: 10px; background-color: #f8f9fa; border-left: 4px solid #dc3545; border-radius: 4px;">
                            <strong>Unique Threat Sources:</strong> {}
                        </li>
                        <li style="margin: 10px 0; padding: 10px; background-color: #f8f9fa; border-left: 4px solid #ffc107; border-radius: 4px;">
                            <strong>Users Affected:</strong> {}
                        </li>
                        <li style="margin: 10px 0; padding: 10px; background-color: #f8f9fa; border-left: 4px solid #fd7e14; border-radius: 4px;">
                            <strong>Activity Log Lockouts:</strong> {}
                        </li>
                    </ul>
                    
                    <h4 style="color: #6f42c1; margin-bottom: 15px; margin-top: 25px;">Account Status Overview</h4>
                    <ul style="list-style: none; padding: 0;">
                        <li style="margin: 10px 0; padding: 10px; background-color: #f8d7da; border-left: 4px solid #dc3545; border-radius: 4px;">
                            <strong>Currently Locked Accounts:</strong> {}
                        </li>
                        <li style="margin: 10px 0; padding: 10px; background-color: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;">
                            <strong>Recently Locked Accounts:</strong> {}
                        </li>
                        <li style="margin: 10px 0; padding: 10px; background-color: #d1ecf1; border-left: 4px solid #17a2b8; border-radius: 4px;">
                            <strong>Accounts Requiring Password Reset:</strong> {}
                        </li>
                        <li style="margin: 10px 0; padding: 10px; background-color: #e2e3e5; border-left: 4px solid #6c757d; border-radius: 4px;">
                            <strong>Accounts with Multiple Failures:</strong> {}
                        </li>
                        <li style="margin: 10px 0; padding: 10px; background-color: #f8f9fa; border-left: 4px solid #28a745; border-radius: 4px;">
                            <strong>Inactive Accounts (90+ days):</strong> {}
                        </li>
                    </ul>
                </div>
                <div class="col-md-6">
                    <h4 style="color: #6f42c1; margin-bottom: 15px;">Audit Summary</h4>
                    <div style="padding: 15px; background-color: #e3f2fd; border-radius: 6px; border-left: 4px solid #2196f3; margin-bottom: 20px;">
                        <p style="margin: 0; font-size: 14px; line-height: 1.6;">
                            <strong>Monitoring Period:</strong> {} days<br>
                            <strong>Days with Incidents:</strong> {} out of {} days<br>
                            <strong>Incident Rate:</strong> {}% of monitoring period<br>
                            <strong>Password Reset Requests:</strong> {}<br>
                            <strong>Security Posture:</strong> {}
                        </p>
                    </div>
                    
                    <h4 style="color: #6f42c1; margin-bottom: 15px;">Risk Assessment</h4>
                    <div style="padding: 15px; background-color: #fff3cd; border-radius: 6px; border-left: 4px solid #ffc107;">
                        <div style="font-size: 14px; color: #856404; line-height: 1.6;">
                            <strong>Overall Risk Level:</strong> 
                            {}
                            <br><br>
                            <strong>Key Risk Indicators:</strong><br>
                            • <strong>Active Threats:</strong> {} unique IP sources<br>
                            • <strong>Account Compromise:</strong> {} locked accounts<br>
                            • <strong>Security Events:</strong> {} total incidents<br>
                            • <strong>User Impact:</strong> {} users affected
                        </div>
                    </div>
                </div>
            </div>
            
            <div style="margin-top: 20px; padding: 15px; background-color: #d4edda; border-left: 4px solid #28a745; border-radius: 4px;">
                <h5 style="color: #155724; margin-top: 0;">Compliance Recommendations</h5>
                <ul style="margin: 0; color: #155724;">
                    <li>Regular security monitoring and incident response procedures are operational</li>
                    <li>Failed authentication attempts are being tracked and analyzed</li>
                    <li>Account lockout policies appear to be functioning (based on lockout events)</li>
                    <li>Password reset capabilities are available and being utilized</li>
                    <li>Consider implementing additional multi-factor authentication for high-risk accounts</li>
                    <li>Review and audit inactive accounts for potential security risks</li>
                    <li>Implement automated alerts for accounts with multiple failed login attempts</li>
                </ul>
            </div>
            
            <div style="margin-top: 20px; padding: 15px; background-color: #e2e3e5; border-left: 4px solid #6c757d; border-radius: 4px;">
                <h5 style="color: #495057; margin-top: 0;">Compliance Status Summary</h5>
                <div style="font-size: 13px; color: #495057;">
                    <p style="margin: 5px 0;"><strong>✓ Security Monitoring:</strong> Active and operational</p>
                    <p style="margin: 5px 0;"><strong>✓ Incident Tracking:</strong> {} security events recorded</p>
                    <p style="margin: 5px 0;"><strong>✓ Account Protection:</strong> {} lockout events detected</p>
                    <p style="margin: 5px 0;"><strong>✓ Password Security:</strong> {} reset requests processed</p>
                    <p style="margin: 5px 0;"><strong>→ Review Required:</strong> {} accounts need attention</p>
                </div>
            </div>
        </div>
    </div>
                    """.format(
                        total_events,
                        unique_threats,
                        users_affected,
                        activity_lockouts,
                        currently_locked,
                        recently_locked,
                        accounts_requiring_reset,
                        accounts_multiple_failures,
                        inactive_accounts,
                        lookback_days,
                        days_with_incidents,
                        lookback_days,
                        incident_rate_str,
                        password_resets,
                        security_posture,
                        "Medium Risk" if total_events > 100 else "Low Risk" if total_events > 0 else "Minimal Risk",
                        unique_threats,
                        currently_locked,
                        total_events,
                        users_affected,
                        total_events,
                        activity_lockouts,
                        password_resets,
                        currently_locked + accounts_requiring_reset + accounts_multiple_failures
                    )
            
        except Exception as e:
            self.print_error("Compliance Audit Report Generation", e)
    
    def generate_remediation_recommendations(self, lookback_days, high_risk_threshold, auto_block_threshold):
        """Generate actionable remediation recommendations"""
        try:
            print """
<div style="border: 2px solid #28a745; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
    <div style="background: linear-gradient(135deg, #c3e6cb 0%, #badbcc 100%); border-bottom: 1px solid #a9d1bc; padding: 15px 20px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 16px; color: #333333;">
        <i class="fa fa-wrench"></i> Security Remediation Recommendations
    </div>
    <div style="padding: 20px;">
        <div class="row">
            <div class="col-md-6">
                <h4 style="color: #155724; margin-bottom: 15px;">Immediate Actions</h4>
                <div style="padding: 15px; background-color: #f8d7da; border-left: 4px solid #dc3545; border-radius: 4px; margin-bottom: 15px;">
                    <h5 style="color: #721c24; margin-top: 0;">Critical Priority</h5>
                    <ul style="margin: 0; color: #721c24;">
                        <li>Block IP addresses exceeding {} failed attempts</li>
                        <li>Force password resets for compromised accounts</li>
                        <li>Enable account lockout after 5 failed attempts</li>
                        <li>Implement immediate rate limiting</li>
                    </ul>
                </div>
                
                <div style="padding: 15px; background-color: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;">
                    <h5 style="color: #856404; margin-top: 0;">High Priority</h5>
                    <ul style="margin: 0; color: #856404;">
                        <li>Review and update firewall rules</li>
                        <li>Enable multi-factor authentication</li>
                        <li>Implement geographic IP restrictions</li>
                        <li>Audit user access permissions</li>
                    </ul>
                </div>
            </div>
            
            <div class="col-md-6">
                <h4 style="color: #155724; margin-bottom: 15px;">Long-term Improvements</h4>
                <div style="padding: 15px; background-color: #d4edda; border-left: 4px solid #28a745; border-radius: 4px; margin-bottom: 15px;">
                    <h5 style="color: #155724; margin-top: 0;">Security Enhancements</h5>
                    <ul style="margin: 0; color: #155724;">
                        <li>Deploy intrusion detection system</li>
                        <li>Implement behavioral analytics</li>
                        <li>Regular security awareness training</li>
                        <li>Automated threat response capabilities</li>
                    </ul>
                </div>
                
                <div style="padding: 15px; background-color: #d1ecf1; border-left: 4px solid #17a2b8; border-radius: 4px;">
                    <h5 style="color: #0c5460; margin-top: 0;">Monitoring & Analytics</h5>
                    <ul style="margin: 0; color: #0c5460;">
                        <li>Real-time security monitoring dashboard</li>
                        <li>Automated alert notifications</li>
                        <li>Regular security assessment reports</li>
                        <li>Threat intelligence integration</li>
                    </ul>
                </div>
            </div>
        </div>
        
        <div style="margin-top: 20px; padding: 20px; background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); border-radius: 6px; border-left: 4px solid #2196f3;">
            <h4 style="color: #1565c0; margin-top: 0;">Implementation Roadmap</h4>
            <div style="color: #1565c0; font-size: 14px;">
                <p><strong>Week 1:</strong> Implement immediate blocking and rate limiting measures</p>
                <p><strong>Week 2-4:</strong> Deploy multi-factor authentication and review access controls</p>
                <p><strong>Month 2:</strong> Enhance monitoring capabilities and automated response</p>
                <p><strong>Month 3+:</strong> Ongoing security improvements and threat intelligence integration</p>
            </div>
        </div>
    </div>
</div>
            """.format(auto_block_threshold)
            
        except Exception as e:
            self.print_error("Remediation Recommendations Generation", e)
    
    def generate_enhanced_login_comparison(self, lookback_days):
        """Generate enhanced login success vs failure comparison"""
        try:
            sql_comparison = """
                DECLARE @StartDate DATETIME = DATEADD(DAY, -{}, GETDATE());
                
                SELECT 
                    'Failed Logins' AS LoginType,
                    COUNT(*) AS TotalAttempts,
                    COUNT(DISTINCT ClientIp) AS UniqueIPs,
                    COUNT(DISTINCT UserId) AS UniqueUsers,
                    COUNT(DISTINCT CAST(ActivityDate AS DATE)) AS ActiveDays
                FROM ActivityLog
                WHERE Activity LIKE '%failed%'
                    AND ActivityDate >= @StartDate
                
                UNION ALL
                
                SELECT 
                    'Successful Logins' AS LoginType,
                    COUNT(*) AS TotalAttempts,
                    COUNT(DISTINCT ClientIp) AS UniqueIPs,
                    COUNT(DISTINCT UserId) AS UniqueUsers,
                    COUNT(DISTINCT CAST(ActivityDate AS DATE)) AS ActiveDays
                FROM ActivityLog
                WHERE (Activity LIKE '%logged in%' OR Activity LIKE '%login%') 
                    AND Activity NOT LIKE '%failed%'
                    AND ActivityDate >= @StartDate;
            """.format(lookback_days)
            
            comparison_data = q.QuerySql(sql_comparison)
            
            if comparison_data and len(comparison_data) > 0:
                failed_data = None
                success_data = None
                
                for row in comparison_data:
                    if row.LoginType == 'Failed Logins':
                        failed_data = row
                    else:
                        success_data = row
                
                print """
<div style="border: 2px solid #20c997; border-radius: 8px; margin: 20px 0; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
    <div style="background: linear-gradient(135deg, #c3e6cb 0%, #badbcc 100%); border-bottom: 1px solid #a9d1bc; padding: 12px 18px; border-radius: 6px 6px 0 0; font-weight: bold; font-size: 16px; color: #333333;">
        <i class="fa fa-chart-pie"></i> Enhanced Login Success vs Failure Analysis
    </div>
    <div style="padding: 20px;">
        <div class="row" style="margin-bottom: 20px;">
            <div class="col-md-6">
                <div style="border: 2px solid #dc3545; background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%); padding: 20px; border-radius: 8px; margin-bottom: 15px;">
                    <div style="font-size: 16px; font-weight: bold; color: #721c24; margin-bottom: 12px;">
                        <i class="fa fa-times-circle"></i> Failed Login Analysis
                    </div>
                    <div style="font-size: 14px; line-height: 1.6; color: #721c24;">
                        <strong>Total Attempts:</strong> {:,}<br>
                        <strong>Unique IP Sources:</strong> {:,}<br>
                        <strong>Affected Users:</strong> {:,}<br>
                        <strong>Active Days:</strong> {} days
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div style="border: 2px solid #28a745; background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%); padding: 20px; border-radius: 8px; margin-bottom: 15px;">
                    <div style="font-size: 16px; font-weight: bold; color: #155724; margin-bottom: 12px;">
                        <i class="fa fa-check-circle"></i> Successful Login Analysis
                    </div>
                    <div style="font-size: 14px; line-height: 1.6; color: #155724;">
                        <strong>Total Attempts:</strong> {:,}<br>
                        <strong>Unique IP Sources:</strong> {:,}<br>
                        <strong>Active Users:</strong> {:,}<br>
                        <strong>Active Days:</strong> {} days
                    </div>
                </div>
            </div>
        </div>
        
        <div style="padding: 18px; background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-radius: 6px; border-left: 4px solid #ffc107;">
            <h4 style="color: #856404; margin-top: 0;">Statistical Analysis</h4>
                """.format(
                    failed_data.TotalAttempts if failed_data else 0,
                    failed_data.UniqueIPs if failed_data else 0,
                    failed_data.UniqueUsers if failed_data else 0,
                    failed_data.ActiveDays if failed_data else 0,
                    success_data.TotalAttempts if success_data else 0,
                    success_data.UniqueIPs if success_data else 0,
                    success_data.UniqueUsers if success_data else 0,
                    success_data.ActiveDays if success_data else 0
                )
                
                if failed_data and success_data and success_data.TotalAttempts > 0:
                    total_attempts = failed_data.TotalAttempts + success_data.TotalAttempts
                    failure_rate = (float(failed_data.TotalAttempts) / total_attempts) * 100
                    success_rate = (float(success_data.TotalAttempts) / total_attempts) * 100
                    
                    print """
            <div style="font-size: 14px; color: #856404; line-height: 1.6;">
                <strong>Overall Authentication Statistics:</strong><br>
                • <strong>Total Authentication Attempts:</strong> {:,}<br>
                • <strong>Success Rate:</strong> {:.1f}% ({:,} successful)<br>
                • <strong>Failure Rate:</strong> {:.1f}% ({:,} failed)<br>
                • <strong>Security Ratio:</strong> {:.1f} failed attempts per successful login<br>
                • <strong>IP Diversity:</strong> {} unique IPs for failures vs {} for successes<br>
                • <strong>User Activity:</strong> {} users with failures vs {} with successes
            </div>
                    """.format(
                        total_attempts,
                        success_rate, success_data.TotalAttempts,
                        failure_rate, failed_data.TotalAttempts,
                        float(failed_data.TotalAttempts) / success_data.TotalAttempts,
                        failed_data.UniqueIPs,
                        success_data.UniqueIPs,
                        failed_data.UniqueUsers,
                        success_data.UniqueUsers
                    )
                else:
                    print "<div style='color: #856404;'>Insufficient data for statistical analysis.</div>"
                
                print """
        </div>
    </div>
</div>
                """
                
        except Exception as e:
            self.print_error("Enhanced Login Comparison Generation", e)
    
    # ::START:: Enhanced Helper Methods
    def analyze_ip_geography(self, ip_address):
        """Analyze IP geography for enhanced threat intelligence"""
        try:
            # Basic IP class analysis since we can't make external API calls
            ip_parts = ip_address.split('.')
            if len(ip_parts) >= 2:
                first_octet = int(ip_parts[0])
                second_octet = int(ip_parts[1])
                
                if first_octet == 10:
                    return '<small style="color: #28a745;">Private: 10.x.x.x</small>'
                elif first_octet == 172 and 16 <= second_octet <= 31:
                    return '<small style="color: #28a745;">Private: 172.16-31.x.x</small>'
                elif first_octet == 192 and second_octet == 168:
                    return '<small style="color: #28a745;">Private: 192.168.x.x</small>'
                elif first_octet == 127:
                    return '<small style="color: #17a2b8;">Localhost: 127.x.x.x</small>'
                else:
                    # Rough geographic indicators based on IP ranges
                    if first_octet >= 1 and first_octet <= 126:
                        return '<small style="color: #6c757d;">Public: Class A range</small>'
                    elif first_octet >= 128 and first_octet <= 191:
                        return '<small style="color: #6c757d;">Public: Class B range</small>'
                    elif first_octet >= 192 and first_octet <= 223:
                        return '<small style="color: #6c757d;">Public: Class C range</small>'
                    else:
                        return '<small style="color: #dc3545;">Special: Reserved range</small>'
            return ''
        except:
            return ''
    
    def calculate_sophistication_level(self, attack_types, total_attempts):
        """Calculate attack sophistication level"""
        if attack_types >= 5 and total_attempts >= 100:
            return "Advanced"
        elif attack_types >= 3 and total_attempts >= 50:
            return "Intermediate"
        elif attack_types >= 2:
            return "Basic"
        else:
            return "Simple"
    
    def get_ip_recommendation(self, risk_level, total_attempts, attack_pattern):
        """Get specific recommendation for IP address"""
        if 'Critical' in risk_level:
            return "IMMEDIATE BLOCKING REQUIRED"
        elif 'High' in risk_level:
            return "Block and monitor closely"
        elif 'Medium' in risk_level:
            return "Increase monitoring frequency"
        elif total_attempts >= 20:
            return "Consider rate limiting"
        else:
            return "Continue standard monitoring"
    
    def get_activity_filter_display_name(self, activity_filter):
        """Convert activity filter to user-friendly display name"""
        filter_map = {
            '%failed%': 'All Failed Activities',
            '%password%': 'Password Related Events',
            '%logged%': 'Failed Login Attempts', 
            '%locked%': 'Account Lockouts',
            'ForgotPassword': 'Password Reset Attempts',
            'User account locked': 'Account Locked Events',
            '%invalid%': 'Invalid Login Attempts',
            '%mobile%': 'Mobile App Failures'
        }
        return filter_map.get(activity_filter, activity_filter)
    
    def extract_username_from_activity(self, activity_text):
        """Extract attempted username/email from activity log text"""
        try:
            if not activity_text:
                return None
                
            activity_lower = activity_text.lower()
            
            # Pattern 1: "failed password #0 by username" or "failed password #0 by email@domain.com"
            if 'failed password' in activity_lower and ' by ' in activity_lower:
                parts = activity_text.split(' by ')
                if len(parts) >= 2:
                    username_part = parts[-1].strip()
                    if ' ' in username_part:
                        username_part = username_part.split()[0]
                    return username_part
            
            # Pattern 2: "forgotpassword(username)" or "forgotpassword(email@domain.com)"
            elif 'forgotpassword' in activity_lower:
                if '(' in activity_text and ')' in activity_text:
                    start = activity_text.find('(')
                    end = activity_text.find(')', start)
                    if start != -1 and end != -1:
                        return activity_text[start+1:end].strip()
            
            # Pattern 3: Look for email addresses anywhere in the text
            elif '@' in activity_text:
                import re
                email_pattern = r'[\w\.-]+@[\w\.-]+'
                email_match = re.search(email_pattern, activity_text)
                if email_match:
                    return email_match.group(0)
            
            return None
        except Exception as e:
            return None
    
    def format_duration(self, seconds):
        """Format duration in seconds to human readable format"""
        if seconds < 60:
            return "{} sec".format(seconds)
        elif seconds < 3600:
            return "{} min".format(seconds // 60)
        else:
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            return "{}h {}m".format(hours, minutes)
    
    def calculate_dotnet_time_span(self, start_time, end_time):
        """Calculate time span between two .NET DateTime objects"""
        try:
            if start_time == end_time:
                return "Instant"
            return "Time span"
        except:
            return "Unknown"
    
    def format_dotnet_datetime(self, dt):
        """Format .NET DateTime object to string"""
        try:
            if dt is None:
                return 'N/A'
            return str(dt)[:19]
        except:
            return 'N/A'
    
    def format_dotnet_datetime_short(self, dt):
        """Format .NET DateTime object to short string"""
        try:
            if dt is None:
                return 'N/A'
            dt_str = str(dt)
            return dt_str[:16]
        except:
            return 'N/A'
    
    def format_dotnet_time(self, dt):
        """Format .NET DateTime object to time only"""
        try:
            if dt is None:
                return 'N/A'
            dt_str = str(dt)
            if ' ' in dt_str:
                return dt_str.split(' ')[1][:8]
            return dt_str[:8]
        except:
            return 'N/A'
    
    def print_error(self, operation, exception):
        """Print formatted error message with expandable details"""
        import random
        error_id = "error_{}".format(random.randint(1000, 9999))
        
        print """
        <div class="alert alert-danger" style="margin: 20px 0;">
            <h4 style="margin-top: 0;">
                <i class="fa fa-exclamation-circle"></i> Error in {}
            </h4>
            <p><strong>Error Message:</strong> {}</p>
            <p>
                <a href="javascript:void(0);" onclick="toggleDetails('{}')" style="color: #721c24; text-decoration: underline;">
                    <i class="fa fa-chevron-down"></i> Click here for technical details
                </a>
            </p>
            <div id="{}" class="tpsecuritydash2025_details" style="margin-top: 15px; background-color: #f5c6cb; padding: 15px; border-radius: 4px;">
                <strong>Stack Trace:</strong>
                <pre style="background-color: #ffffff; padding: 10px; border: 1px solid #f5c6cb; border-radius: 4px; overflow-x: auto;">""".format(
            operation, 
            str(exception), 
            error_id,
            error_id
        )
        
        traceback.print_exc()
        
        print """</pre>
            </div>
        </div>
        """

# Execute main
if __name__ == "__main__":
    main()
else:
    main()
