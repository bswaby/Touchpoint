# TouchPoint Involvement Sync Tool
# This tool synchronizes members and their subgroup assignments between a primary and secondary involvement
# It supports both manual and automatic synchronization using TouchPoint's ExtraValueOrg system
#
# FEATURES:
# - One-way mirror sync (adds AND removes members/subgroups from Primary --> Secondary)
# - Web-based configuration interface
# - Manual and automatic sync options
# - Full subgroup/member tag synchronization
# - Detailed sync reporting
# - Error handling and logging
#
# USE CASES:
# - Backup classes (sync main class to backup class)
# - Sync Schedule Involvement to resolve check-in issues
# - Multi-campus involvement mirroring
# - Any scenario where you need to keep two involvements identical
#
# HOW IT WORKS:
# 1. Configure sync relationships through the web interface
# 2. Enable auto-sync for scheduled synchronization
# 3. Sync runs automatically or manually as needed
# 4. All changes (adds/removes) are tracked and reported

#--Upload Instructions Start--
# To upload code to Touchpoint, use the following steps:
# 1. Click Admin > Advanced > Special Content > Python
# 2. Click New Python Script File
# 3. Name the Python "InvolvementSync" (or your preferred name)
# 4. Paste all this code
# 5. Test the script
# 6. Optionally add to menu for easy access
#
# For Automatic Sync:
# 1. Set up your sync configurations with "Auto Sync" enabled
# 2. Add these lines to your ScheduledTasks or MorningBatch script:
#    
#    # Involvement Sync - Daily at 6 AM
#    if model.DayOfWeek in [0,1,2,3,4,5,6] and model.ScheduledTime == '0600': 
#        model.Data.ScriptSyncAll = 1
#        print(model.CallScript("InvolvementSync"))  # Update to your script name
#
# 3. Adjust the schedule and script name as needed
#--Upload Instructions End--

import traceback
import re

model.Header = "Involvement Sync Manager"

class InvolvementSyncManager:
    """Manages synchronization between primary and secondary involvements"""
    
    def __init__(self):
        self.sync_enabled_field = "SyncEnabled"
        self.sync_target_field = "SyncTargetOrgId"
        self.sync_auto_field = "AutoSync"
        self.loading_html = '''
        <div id="loading" style="display:none; text-align:center; margin:20px;">
            <div style="font-size:18px; color:#666;">
                <i class="fa fa-spinner fa-spin"></i> Processing...
            </div>
        </div>
        '''
    
    def get_involvement_name(self, org_id):
        """Get the name of an involvement by ID"""
        try:
            sql = "SELECT OrganizationName FROM Organizations WHERE OrganizationId = " + str(org_id)
            result = q.QuerySqlTop1(sql)
            return result.OrganizationName if result else "Unknown"
        except:
            return "Unknown"
    
    def get_involvement_details(self, org_id):
        """Get detailed information about an involvement"""
        try:
            # Simplified query - just get the basic org info we need
            sql = '''
            SELECT o.OrganizationId, o.OrganizationName, o.MemberCount
            FROM Organizations o
            WHERE o.OrganizationId = ''' + str(org_id)
            return q.QuerySqlTop1(sql)
        except Exception as e:
            print "<!-- DEBUG ERROR in get_involvement_details for org " + str(org_id) + ": " + str(e) + " -->"
            return None

    def verify_organization_exists(self, org_id):
        """Simple check if organization exists without complex joins"""
        try:
            sql = "SELECT OrganizationId FROM Organizations WHERE OrganizationId = " + str(org_id)
            result = q.QuerySqlTop1(sql)
            return result is not None
        except Exception as e:
            print "<!-- DEBUG ERROR in verify_organization_exists for org " + str(org_id) + ": " + str(e) + " -->"
            return False
    
    def get_sync_enabled_involvements(self):
        """Get all involvements that have sync enabled"""
        try:
            # Simple approach - just look for the SyncEnabled field with True value
            sql = '''
            SELECT DISTINCT o.OrganizationId, o.OrganizationName
            FROM Organizations o
            INNER JOIN OrganizationExtra oe ON o.OrganizationId = oe.OrganizationId
            WHERE oe.Field = \'''' + self.sync_enabled_field + '''\' 
            AND (oe.IntValue = 1 OR oe.BitValue = 'True' OR oe.Data = 'True' OR oe.Data = '1')
            ORDER BY o.OrganizationName
            '''
            result = q.QuerySql(sql)
            
            return result
        except Exception as e:
            print "<!-- DEBUG ERROR in get_sync_enabled_involvements: " + str(e) + " -->"
            return []
    
    def get_sync_settings(self, org_id):
        """Get sync settings for an involvement"""
        try:
            settings = {}
            
            # Get sync enabled status
            settings['enabled'] = False
            enabled_sql = '''
            SELECT Data, BitValue, IntValue FROM OrganizationExtra 
            WHERE OrganizationId = ''' + str(org_id) + ''' AND Field = \'''' + self.sync_enabled_field + '''\'
            '''
            enabled_result = q.QuerySqlTop1(enabled_sql)
            if enabled_result:
                # Check multiple ways the value might be stored
                data_value = str(enabled_result.Data).upper() if enabled_result.Data else ""
                bit_value = str(enabled_result.BitValue).upper() if enabled_result.BitValue else ""
                
                if (enabled_result.IntValue == 1 or 
                    data_value in ['TRUE', '1', 'YES', 'ENABLED'] or
                    bit_value in ['TRUE', '1', 'YES', 'ENABLED']):
                    settings['enabled'] = True
            
            # Get target organization ID
            target_sql = '''
            SELECT Data, IntValue FROM OrganizationExtra
            WHERE OrganizationId = ''' + str(org_id) + ''' AND Field = \'''' + self.sync_target_field + '''\'
            '''
            target_result = q.QuerySqlTop1(target_sql)
            if target_result:
                # Debug output
                print "<!-- DEBUG: For org " + str(org_id) + ", IntValue=" + str(target_result.IntValue) + ", Data=" + str(target_result.Data) + " -->"

                # Try IntValue first, then Data
                if target_result.IntValue and target_result.IntValue > 0:
                    settings['target_org_id'] = target_result.IntValue
                elif target_result.Data and str(target_result.Data).strip().isdigit():
                    settings['target_org_id'] = int(str(target_result.Data).strip())
                else:
                    settings['target_org_id'] = None
            else:
                print "<!-- DEBUG: No target org extra value found for org " + str(org_id) + " -->"
                settings['target_org_id'] = None
            
            # Get auto sync status
            auto_sql = '''
            SELECT Data, BitValue, IntValue FROM OrganizationExtra 
            WHERE OrganizationId = ''' + str(org_id) + ''' AND Field = \'''' + self.sync_auto_field + '''\'
            '''
            auto_result = q.QuerySqlTop1(auto_sql)
            settings['auto_sync'] = False
            if auto_result:
                auto_data_value = str(auto_result.Data).upper() if auto_result.Data else ""
                auto_bit_value = str(auto_result.BitValue).upper() if auto_result.BitValue else ""
                
                if (auto_result.IntValue == 1 or 
                    auto_data_value in ['TRUE', '1', 'YES', 'ENABLED'] or
                    auto_bit_value in ['TRUE', '1', 'YES', 'ENABLED']):
                    settings['auto_sync'] = True
            
            return settings
        except Exception as e:
            print "<!-- DEBUG ERROR in get_sync_settings for org " + str(org_id) + ": " + str(e) + " -->"
            return {'enabled': False, 'target_org_id': None, 'auto_sync': False}
    
    def set_sync_settings(self, org_id, target_org_id, auto_sync=False):
        """Enable sync for an involvement"""
        try:
            print "<!-- DEBUG: Setting sync settings for org " + str(org_id) + " with target " + str(target_org_id) + " -->"

            # Enable sync
            model.AddExtraValueBoolOrg(org_id, self.sync_enabled_field, True)

            # Set target organization - ensure it's an integer
            model.AddExtraValueIntOrg(org_id, self.sync_target_field, int(target_org_id))

            # Set auto sync
            model.AddExtraValueBoolOrg(org_id, self.sync_auto_field, auto_sync)

            # Verify it was saved correctly
            test_settings = self.get_sync_settings(org_id)
            if test_settings['target_org_id'] != int(target_org_id):
                print "<!-- WARNING: Target org ID mismatch after save. Expected " + str(target_org_id) + ", got " + str(test_settings['target_org_id']) + " -->"

            return True
        except Exception as e:
            print "<p style='color:red;'>Error setting sync settings: " + str(e) + "</p>"
            return False
    
    def disable_sync(self, org_id):
        """Disable sync for an involvement"""
        try:
            model.AddExtraValueBoolOrg(org_id, self.sync_enabled_field, False)
            model.AddExtraValueBoolOrg(org_id, self.sync_auto_field, False)
            return True
        except Exception as e:
            print "<p style='color:red;'>Error disabling sync: " + str(e) + "</p>"
            return False
    
    def create_duplicate_involvement(self, source_org_id, new_name=None):
        """Create a duplicate involvement based on source"""
        try:
            # First verify the source organization exists
            source_sql = "SELECT OrganizationId, OrganizationName FROM Organizations WHERE OrganizationId = " + str(source_org_id)
            source_check = q.QuerySqlTop1(source_sql)
            if not source_check:
                raise Exception("Source involvement with ID " + str(source_org_id) + " does not exist")
            
            if not new_name:
                new_name = source_check.OrganizationName + " - Duplicate"
            
            # Try the AddOrganization method with template (preferred method)
            try:
                new_org_id = model.AddOrganization(new_name, source_org_id, True)
                if new_org_id and new_org_id > 0:
                    return new_org_id
            except Exception as e:
                print "<p style='color:orange;'>Template method failed: " + str(e) + "</p>"
            
            # If template method fails, try basic creation
            # Note: We skip trying to use Program/Division since they may not exist
            print "<!-- Template method failed, will try default creation -->"
            
            # If template method fails, try with just the name
            try:
                # Try the simplest form - just organization name
                new_org_id = model.AddOrganization(new_name)
                if new_org_id and new_org_id > 0:
                    return new_org_id
            except Exception as e:
                print "<p style='color:orange;'>Simple creation failed: " + str(e) + "</p>"

            raise Exception("Failed to create new involvement. Please use an existing involvement ID instead.")
            
        except Exception as e:
            print "<p style='color:red;'>Error creating duplicate involvement: " + str(e) + "</p>"
            return None
    
    def get_members_with_subgroups(self, org_id):
        """Get all members of an involvement with their subgroup assignments"""
        try:
            sql = '''
            SELECT om.PeopleId, p.Name, om.OrganizationId,
                   STUFF((
                       SELECT ', ' + mt.Name
                       FROM OrgMemMemTags ommt
                       INNER JOIN MemberTags mt ON mt.Id = ommt.MemberTagId
                       WHERE ommt.OrgId = om.OrganizationId 
                       AND ommt.PeopleId = om.PeopleId
                       FOR XML PATH('')
                   ), 1, 2, '') as SubGroups
            FROM OrganizationMembers om
            INNER JOIN People p ON om.PeopleId = p.PeopleId
            WHERE om.OrganizationId = ''' + str(org_id) + '''
            AND om.InactiveDate IS NULL
            ORDER BY p.Name
            '''
            return q.QuerySql(sql)
        except Exception as e:
            print "<p style='color:red;'>Error getting members: " + str(e) + "</p>"
            return []
    
    def sync_members(self, source_org_id, target_org_id):
        """Sync members and subgroups from source to target involvement"""
        try:
            results = {
                'members_added': 0,
                'members_updated': 0,
                'members_removed': 0,
                'subgroups_synced': 0,
                'subgroups_removed': 0,
                'errors': []
            }
            
            # Get source members with subgroups
            source_members = self.get_members_with_subgroups(source_org_id)
            source_member_ids = [member.PeopleId for member in source_members]
            
            # Get current target members with their subgroups
            target_members = self.get_members_with_subgroups(target_org_id)
            target_member_ids = [member.PeopleId for member in target_members]
            
            # Create lookup dictionaries for easier processing
            source_member_subgroups = {}
            for member in source_members:
                subgroups = set()
                if member.SubGroups and member.SubGroups.strip():
                    subgroups = set([sg.strip() for sg in member.SubGroups.split(',') if sg.strip()])
                source_member_subgroups[member.PeopleId] = subgroups
            
            target_member_subgroups = {}
            for member in target_members:
                subgroups = set()
                if member.SubGroups and member.SubGroups.strip():
                    subgroups = set([sg.strip() for sg in member.SubGroups.split(',') if sg.strip()])
                target_member_subgroups[member.PeopleId] = subgroups
            
            # 1. ADD NEW MEMBERS
            for member in source_members:
                try:
                    people_id = member.PeopleId
                    
                    if people_id not in target_member_ids:
                        model.AddMemberToOrg(people_id, target_org_id)
                        results['members_added'] += 1
                    else:
                        results['members_updated'] += 1
                        
                except Exception as e:
                    error_msg = "Error adding member {0}: {1}".format(member.Name, str(e))
                    results['errors'].append(error_msg)
            
            # 2. REMOVE MEMBERS NO LONGER IN SOURCE
            for member in target_members:
                try:
                    people_id = member.PeopleId
                    
                    if people_id not in source_member_ids:
                        model.DropOrgMember(people_id, target_org_id)
                        results['members_removed'] += 1
                        
                except Exception as e:
                    error_msg = "Error removing member {0}: {1}".format(member.Name, str(e))
                    results['errors'].append(error_msg)
            
            # 3. SYNC SUBGROUPS FOR ALL CURRENT MEMBERS
            for people_id in source_member_ids:
                try:
                    source_subgroups = source_member_subgroups.get(people_id, set())
                    target_subgroups = target_member_subgroups.get(people_id, set())
                    
                    # Add new subgroups
                    subgroups_to_add = source_subgroups - target_subgroups
                    for subgroup_name in subgroups_to_add:
                        try:
                            model.AddSubGroup(people_id, target_org_id, subgroup_name)
                            results['subgroups_synced'] += 1
                        except Exception as sg_error:
                            member_name = next((m.Name for m in source_members if m.PeopleId == people_id), "Unknown")
                            error_msg = "Error adding subgroup '{0}' for {1}: {2}".format(subgroup_name, member_name, str(sg_error))
                            results['errors'].append(error_msg)
                    
                    # Remove subgroups no longer in source
                    subgroups_to_remove = target_subgroups - source_subgroups
                    for subgroup_name in subgroups_to_remove:
                        try:
                            model.RemoveSubGroup(people_id, target_org_id, subgroup_name)
                            results['subgroups_removed'] += 1
                        except Exception as sg_error:
                            member_name = next((m.Name for m in source_members if m.PeopleId == people_id), "Unknown")
                            error_msg = "Error removing subgroup '{0}' for {1}: {2}".format(subgroup_name, member_name, str(sg_error))
                            results['errors'].append(error_msg)
                            
                except Exception as e:
                    member_name = next((m.Name for m in source_members if m.PeopleId == people_id), "Unknown")
                    error_msg = "Error syncing subgroups for {0}: {1}".format(member_name, str(e))
                    results['errors'].append(error_msg)
            
            return results
        except Exception as e:
            return {'members_added': 0, 'members_updated': 0, 'members_removed': 0, 
                   'subgroups_synced': 0, 'subgroups_removed': 0,
                   'errors': ["Error during sync: " + str(e)]}
    
    def render_sync_results(self, results):
        """Render sync results in HTML"""
        html = "<div class='alert alert-success'>"
        html += "<h4>Sync Complete!</h4>"
        html += "<ul>"
        html += "<li>Members Added: " + str(results['members_added']) + "</li>"
        html += "<li>Members Updated: " + str(results['members_updated']) + "</li>"
        html += "<li>Members Removed: " + str(results['members_removed']) + "</li>"
        html += "<li>Subgroups Added: " + str(results['subgroups_synced']) + "</li>"
        html += "<li>Subgroups Removed: " + str(results['subgroups_removed']) + "</li>"
        html += "</ul>"
        
        if results['errors']:
            html += "<h5 style='color:orange;'>Errors:</h5><ul>"
            for error in results['errors']:
                html += "<li style='color:red;'>" + error + "</li>"
            html += "</ul>"
        
        html += "</div>"
        return html

def get_common_css():
    """Get the common CSS used across all forms"""
    return '''
    <style>
        #involvement-sync-container {
            max-width: 1000px;
            margin: 20px auto;
            padding: 20px;
            font-family: Arial, sans-serif;
            background: #fff;
            border: 1px solid #ddd;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        #involvement-sync-container h2, #involvement-sync-container h3 {
            color: #333;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #337ab7;
            font-weight: 600;
        }
        #involvement-sync-container .sync-form-group {
            margin-bottom: 20px;
        }
        #involvement-sync-container .sync-label {
            display: block;
            margin-bottom: 8px;
            font-weight: bold;
            color: #333;
            font-size: 14px;
        }
        #involvement-sync-container .sync-input {
            width: 100%;
            padding: 10px 12px;
            border: 2px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
            transition: border-color 0.3s ease;
        }
        #involvement-sync-container .sync-input:focus {
            border-color: #337ab7;
            outline: none;
            box-shadow: 0 0 5px rgba(51,122,183,0.3);
        }
        #involvement-sync-container .sync-help {
            display: block;
            margin-top: 5px;
            color: #666;
            font-size: 12px;
        }
        #involvement-sync-container .sync-radio-group {
            margin: 10px 0;
        }
        #involvement-sync-container .sync-radio {
            margin: 8px 0;
        }
        #involvement-sync-container .sync-radio label {
            margin-left: 8px;
            font-weight: normal;
            cursor: pointer;
        }
        #involvement-sync-container .sync-checkbox {
            margin: 15px 0;
        }
        #involvement-sync-container .sync-checkbox label {
            margin-left: 8px;
            font-weight: normal;
            cursor: pointer;
        }
        #involvement-sync-container .sync-btn {
            display: inline-block;
            padding: 12px 24px;
            margin: 5px 10px 5px 0;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            font-weight: bold;
            text-decoration: none;
            transition: background-color 0.3s ease;
        }
        #involvement-sync-container .sync-btn-primary {
            background-color: #337ab7;
            color: white;
        }
        #involvement-sync-container .sync-btn-primary:hover {
            background-color: #286090;
        }
        #involvement-sync-container .sync-btn-success {
            background-color: #5cb85c;
            color: white;
        }
        #involvement-sync-container .sync-btn-success:hover {
            background-color: #449d44;
        }
        #involvement-sync-container .sync-btn-default {
            background-color: #f8f8f8;
            color: #333;
            border: 1px solid #ccc;
        }
        #involvement-sync-container .sync-btn-default:hover {
            background-color: #e8e8e8;
        }
        #involvement-sync-container .sync-alert {
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
        }
        #involvement-sync-container .sync-alert-success {
            border: 1px solid #d6e9c6;
            background-color: #dff0d8;
            color: #3c763d;
        }
        #involvement-sync-container .sync-alert-warning {
            border: 1px solid #faebcc;
            background-color: #fcf8e3;
            color: #8a6d3b;
        }
        #involvement-sync-container .sync-alert-danger {
            border: 1px solid #ebccd1;
            background-color: #f2dede;
            color: #a94442;
        }
        #involvement-sync-container .sync-panel {
            border: 1px solid #ddd;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        #involvement-sync-container .sync-panel-heading {
            padding: 10px 15px;
            background-color: #f5f5f5;
            border-bottom: 1px solid #ddd;
            border-radius: 4px 4px 0 0;
        }
        #involvement-sync-container .sync-panel-body {
            padding: 15px;
        }
        #involvement-sync-container .sync-hidden {
            display: none;
        }
        #involvement-sync-container .sync-radio-option {
            padding: 10px;
            margin: 5px 0;
            border: 1px solid #ddd;
            border-radius: 4px;
            background-color: #f9f9f9;
        }
        #involvement-sync-container .sync-radio-option:hover {
            background-color: #e9e9e9;
        }
        #involvement-sync-container .sync-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }
        #involvement-sync-container .sync-table th,
        #involvement-sync-container .sync-table td {
            padding: 8px 12px;
            border: 1px solid #ddd;
            text-align: left;
        }
        #involvement-sync-container .sync-table th {
            background-color: #f5f5f5;
            font-weight: bold;
        }
    </style>
    '''

def show_loading_script():
    """JavaScript to show/hide loading indicator"""
    return '''
    <script>
        function showLoading() {
            var loadingDiv = document.getElementById('loading');
            if (loadingDiv) {
                loadingDiv.style.display = 'block';
            }
            var forms = document.getElementsByTagName('form');
            for(var i = 0; i < forms.length; i++) {
                var buttons = forms[i].getElementsByTagName('button');
                for(var j = 0; j < buttons.length; j++) {
                    buttons[j].disabled = true;
                }
            }
        }
        
        // Helper function to get the correct form submission URL
        function getPyScriptAddress() {
            let path = window.location.pathname;
            return path.replace("/PyScript/", "/PyScriptForm/");
        }
    </script>
    '''

def render_main_menu():
    """Render the main menu interface"""
    sync_manager = InvolvementSyncManager()
    sync_enabled_orgs = sync_manager.get_sync_enabled_involvements()
    
    html = get_common_css() + '''
    <div id="involvement-sync-container">
        <h2><i class="fa fa-sync"></i> Involvement Sync Manager<svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
                  </svg></h2>
        <p style="font-size: 16px; color: #666; margin-bottom: 30px;">Synchronize members and subgroups from primary --> secondary involvement.</p>
        
        <div style="display: flex; gap: 20px; margin-bottom: 30px;">
            <div style="flex: 1;">
                <div class="sync-panel">
                    <div class="sync-panel-heading">
                        <h4><i class="fa fa-cog"></i> Setup New Sync</h4>
                    </div>
                    <div class="sync-panel-body">
                        <p>Configure synchronization between a primary and secondary involvement.</p>
                        <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading()">
                            <input type="hidden" name="action" value="setup_form">
                            <button type="submit" class="sync-btn sync-btn-primary">
                                <i class="fa fa-plus"></i> Setup New Sync
                            </button>
                        </form>
                    </div>
                </div>
            </div>
            
            <div style="flex: 1;">
                <div class="sync-panel">
                    <div class="sync-panel-heading">
                        <h4><i class="fa fa-play"></i> Manual Sync</h4>
                    </div>
                    <div class="sync-panel-body">
                        <p>Manually synchronize configured involvements.</p>
                        <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading()">
                            <input type="hidden" name="action" value="manual_sync_form">
                            <button type="submit" class="sync-btn sync-btn-success">
                                <i class="fa fa-sync"></i> Manual Sync
                            </button>
                        </form>
                    </div>
                </div>
            </div>
        </div>
    '''
    
    # Show configured sync relationships
    if sync_enabled_orgs:
        html += '''
        <div class="sync-panel">
            <div class="sync-panel-heading">
                <h4><i class="fa fa-list"></i> Configured Sync Relationships</h4>
            </div>
            <div class="sync-panel-body">
                <table class="sync-table">
                    <thead>
                        <tr>
                            <th>Primary Involvement</th>
                            <th>Target Involvement</th>
                            <th>Auto Sync</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
        '''
        
        for org in sync_enabled_orgs:
            settings = sync_manager.get_sync_settings(org.OrganizationId)
            target_name = sync_manager.get_involvement_name(settings['target_org_id']) if settings['target_org_id'] else "Not Set"
            auto_status = "Yes" if settings['auto_sync'] else "No"
            
            html += '''
                        <tr>
                            <td><strong>''' + org.OrganizationName + '''</strong></td>
                            <td>''' + target_name + '''</td>
                            <td><span style="padding: 2px 8px; border-radius: 3px; background-color: ''' + ("#5cb85c" if settings['auto_sync'] else "#999") + '''; color: white; font-size: 12px;">''' + auto_status + '''</span></td>
                            <td>
                                <form method="post" action="" style="display:inline;" onsubmit="this.action = getPyScriptAddress(); showLoading()">
                                    <input type="hidden" name="action" value="sync_now">
                                    <input type="hidden" name="source_org_id" value="''' + str(org.OrganizationId) + '''">
                                    <button type="submit" class="sync-btn sync-btn-primary" style="padding: 6px 12px; font-size: 12px;">
                                        <i class="fa fa-sync"></i> Sync Now
                                    </button>
                                </form>
                                <form method="post" action="" style="display:inline; margin-left:5px;" onsubmit="this.action = getPyScriptAddress(); showLoading()">
                                    <input type="hidden" name="action" value="edit_sync">
                                    <input type="hidden" name="source_org_id" value="''' + str(org.OrganizationId) + '''">
                                    <button type="submit" class="sync-btn sync-btn-default" style="padding: 6px 12px; font-size: 12px;">
                                        <i class="fa fa-edit"></i> Edit
                                    </button>
                                </form>
                            </td>
                        </tr>
            '''
        
        html += '''
                    </tbody>
                </table>
            </div>
        </div>
        '''
    else:
        html += '''
        <div class="sync-alert sync-alert-warning">
            <h4>No Sync Configurations Found</h4>
            <p>You haven't set up any sync configurations yet. Use the "Setup New Sync" button above to get started.</p>
        </div>
        '''
    
    html += '''
    </div>
    ''' + sync_manager.loading_html
    
    return html

def render_setup_form(org_id=None):
    """Render the setup form for configuring sync"""
    sync_manager = InvolvementSyncManager()
    
    # Get existing settings if editing
    settings = {}
    org_name = ""
    if org_id:
        settings = sync_manager.get_sync_settings(org_id)
        org_name = sync_manager.get_involvement_name(org_id)
    
    html = get_common_css() + '''
    <div id="involvement-sync-container">
        <h3><i class="fa fa-cog"></i> ''' + ("Edit" if org_id else "Setup") + ''' Involvement Sync</h3>
        
        <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading()">
            <input type="hidden" name="action" value="''' + ("update_sync" if org_id else "create_sync") + '''">
    '''
    
    if org_id:
        html += '<input type="hidden" name="source_org_id" value="' + str(org_id) + '">'
        html += '<div class="sync-alert sync-alert-success"><strong>Editing sync for:</strong> ' + org_name + '</div>'
    else:
        html += '''
            <div class="sync-form-group">
                <label for="source_org_id" class="sync-label">Primary Involvement ID:</label>
                <input type="number" id="source_org_id" name="source_org_id" class="sync-input" required 
                       placeholder="Enter the ID of the primary involvement">
                <small class="sync-help">This is the involvement that will be the source of the sync.</small>
            </div>
        '''
    
    html += '''
            <div class="sync-form-group">
                <label class="sync-label">Target Involvement:</label>
                <div class="sync-radio-group">
                    <div class="sync-radio">
                        <input type="radio" id="target_existing" name="target_type" value="existing" checked onchange="toggleSyncTargetOptions()">
                        <label for="target_existing">Sync to existing involvement</label>
                    </div>
                    <div class="sync-radio">
                        <input type="radio" id="target_create" name="target_type" value="create" onchange="toggleSyncTargetOptions()">
                        <label for="target_create">Create new involvement</label>
                    </div>
                </div>
            </div>
            
            <div id="sync_existing_target" class="sync-form-group">
                <label for="target_org_id" class="sync-label">Target Involvement ID:</label>
                <input type="number" id="target_org_id" name="target_org_id" class="sync-input"
                       value="''' + str(settings.get('target_org_id', '')) + '''"
                       placeholder="Enter the ID of the target involvement">
            </div>
            
            <div id="sync_new_target" class="sync-form-group sync-hidden">
                <label for="new_org_name" class="sync-label">New Involvement Name:</label>
                <input type="text" id="new_org_name" name="new_org_name" class="sync-input"
                       placeholder="Leave blank to auto-generate name">
                <small class="sync-help">If left blank, will append " - Duplicate" to the primary involvement name.</small>
            </div>
            
            <div class="sync-form-group">
                <div class="sync-checkbox">
                    <input type="checkbox" id="auto_sync" name="auto_sync" value="1" ''' + ("checked" if settings.get('auto_sync') else "") + '''>
                    <label for="auto_sync">Enable automatic synchronization</label>
                </div>
                <small class="sync-help">When enabled, sync will happen automatically during scheduled processes.</small>
            </div>
            
            <div class="sync-form-group">
                <button type="submit" class="sync-btn sync-btn-primary">
                    <i class="fa fa-save"></i> ''' + ("Update" if org_id else "Create") + ''' Sync Configuration
                </button>
                <button type="button" onclick="history.back()" class="sync-btn sync-btn-default">
                    <i class="fa fa-arrow-left"></i> Back to Menu
                </button>
            </div>
        </form>
    </div>
    
    <script>
        function toggleSyncTargetOptions() {
            var existing = document.getElementById('target_existing').checked;
            var existingDiv = document.getElementById('sync_existing_target');
            var newDiv = document.getElementById('sync_new_target');
            
            if (existing) {
                existingDiv.style.display = 'block';
                newDiv.style.display = 'none';
                document.getElementById('new_org_name').value = '';
            } else {
                existingDiv.style.display = 'none';
                newDiv.style.display = 'block';
                document.getElementById('target_org_id').value = '';
            }
        }
    </script>
    ''' + sync_manager.loading_html
    
    return html

def render_manual_sync_form():
    """Render form for manual sync selection"""
    sync_manager = InvolvementSyncManager()
    sync_enabled_orgs = sync_manager.get_sync_enabled_involvements()
    
    html = get_common_css() + '''
    <div id="involvement-sync-container">
        <h3><i class="fa fa-sync"></i> Manual Sync</h3>
    '''
    
    if not sync_enabled_orgs:
        html += '''
        <div class="sync-alert sync-alert-warning">
            <h4>No Sync Configurations Found</h4>
            <p>You need to set up at least one sync configuration before you can perform manual sync.</p>
            <button onclick="history.back()" class="sync-btn sync-btn-primary">
                <i class="fa fa-arrow-left"></i> Back to Menu
            </button>
        </div>
        </div>
        '''
        return html
    
    html += '''
        <p>Select which involvement sync to perform:</p>
        
        <!-- Sync All Button -->
        <div class="sync-panel">
            <div class="sync-panel-body">
                <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading()">
                    <input type="hidden" name="action" value="sync_all_manual">
                    <button type="submit" class="sync-btn sync-btn-success" style="width: 100%; padding: 15px; font-size: 18px;">
                        <i class="fa fa-sync"></i> Sync All Configured Involvements
                    </button>
                </form>
            </div>
        </div>
        
        <h4>Or Sync Individual Involvement:</h4>
        <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading()">
            <input type="hidden" name="action" value="perform_manual_sync">
            
            <div class="sync-form-group">
    '''
    
    for org in sync_enabled_orgs:
        settings = sync_manager.get_sync_settings(org.OrganizationId)
        target_name = sync_manager.get_involvement_name(settings['target_org_id']) if settings['target_org_id'] else "Not Set"
        
        html += '''
                <div class="sync-radio-option">
                    <label style="display: block; cursor: pointer;">
                        <input type="radio" name="source_org_id" value="''' + str(org.OrganizationId) + '''" style="margin-right: 10px;">
                        <strong>''' + org.OrganizationName + '''</strong> → ''' + target_name + '''
                    </label>
                </div>
        '''
    
    html += '''
            </div>
            
            <div class="sync-form-group">
                <button type="submit" class="sync-btn sync-btn-primary">
                    <i class="fa fa-sync"></i> Sync Selected
                </button>
                <button type="button" onclick="history.back()" class="sync-btn sync-btn-default">
                    <i class="fa fa-arrow-left"></i> Back to Menu
                </button>
            </div>
        </form>
    </div>
    ''' + sync_manager.loading_html
    
    return html

# Main execution logic
sync_manager = InvolvementSyncManager()

# Check for automated sync URL parameter
script_sync_all = ""
try:
    script_sync_all = str(model.Data.ScriptSyncAll) if hasattr(model.Data, 'ScriptSyncAll') else ""
except:
    script_sync_all = ""

# Handle automated sync
if script_sync_all == "1":
    # Perform automated sync for all auto-enabled involvements
    sync_enabled_orgs = sync_manager.get_sync_enabled_involvements()
    results_summary = []
    
    for org in sync_enabled_orgs:
        settings = sync_manager.get_sync_settings(org.OrganizationId)
        if settings['auto_sync'] and settings['target_org_id']:
            try:
                results = sync_manager.sync_members(org.OrganizationId, settings['target_org_id'])
                source_name = sync_manager.get_involvement_name(org.OrganizationId)
                target_name = sync_manager.get_involvement_name(settings['target_org_id'])
                
                results_summary.append({
                    'source': source_name,
                    'target': target_name,
                    'success': True,
                    'results': results
                })
            except Exception as e:
                results_summary.append({
                    'source': org.OrganizationName,
                    'target': 'Unknown',
                    'success': False,
                    'error': str(e)
                })
    
    # Output simple results for automated process
    print "Automated Sync Results:"
    for result in results_summary:
        if result['success']:
            print "SUCCESS: {0} -> {1} | Added: {2}, Updated: {3}, Removed: {4}, Subgroups Added: {5}, Subgroups Removed: {6}".format(
                result['source'], result['target'], 
                result['results']['members_added'],
                result['results']['members_updated'],
                result['results']['members_removed'],
                result['results']['subgroups_synced'],
                result['results']['subgroups_removed']
            )
        else:
            print "ERROR: {0} -> {1} | {2}".format(
                result['source'], result['target'], result['error']
            )

else:
    # Handle interactive form submissions
    try:
        # Handle form submissions for interactive use
        action = ""
        try:
            action = str(model.Data.action) if hasattr(model.Data, 'action') else ""
        except:
            action = ""
        
        if action == "setup_form":
            print show_loading_script()
            print render_setup_form()
            
        elif action == "manual_sync_form":
            print show_loading_script()
            print render_manual_sync_form()
            
        elif action == "create_sync":
            print show_loading_script()
            
            try:
                source_org_id = int(str(model.Data.source_org_id))
                target_type = str(model.Data.target_type)
                auto_sync = hasattr(model.Data, 'auto_sync') and str(model.Data.auto_sync) == "1"
                
                target_org_id = None
                
                if target_type == "existing":
                    target_org_id = int(str(model.Data.target_org_id))
                    # Verify target org exists with simple check first
                    if not sync_manager.verify_organization_exists(target_org_id):
                        raise Exception("Target involvement ID " + str(target_org_id) + " does not exist in the database")
                    # Try to get details (may fail due to missing division/program)
                    target_details = sync_manager.get_involvement_details(target_org_id)
                    if not target_details:
                        print "<!-- DEBUG: Organization " + str(target_org_id) + " exists but could not get details (possibly missing Division/Program) -->"
                else:
                    # Create new involvement
                    new_name = str(model.Data.new_org_name) if hasattr(model.Data, 'new_org_name') and str(model.Data.new_org_name).strip() else None
                    target_org_id = sync_manager.create_duplicate_involvement(source_org_id, new_name)
                    
                    if not target_org_id:
                        raise Exception("Failed to create new involvement. Please try using an existing involvement instead.")
                
                # Set up sync configuration
                if sync_manager.set_sync_settings(source_org_id, target_org_id, auto_sync):
                    source_name = sync_manager.get_involvement_name(source_org_id)
                    target_name = sync_manager.get_involvement_name(target_org_id)
                    
                    print get_common_css() + '''
                    <div id="involvement-sync-container">
                        <div class="sync-alert sync-alert-success">
                            <h4>Sync Configuration Created Successfully!</h4>
                            <p><strong>Primary:</strong> ''' + source_name + ''' (ID: ''' + str(source_org_id) + ''')</p>
                            <p><strong>Target:</strong> ''' + target_name + ''' (ID: ''' + str(target_org_id) + ''')</p>
                            <p><strong>Auto Sync:</strong> ''' + ("Enabled" if auto_sync else "Disabled") + '''</p>
                            
                            <div style="margin-top:20px;">
                                <form method="post" action="" style="display:inline;" onsubmit="this.action = getPyScriptAddress(); showLoading()">
                                    <input type="hidden" name="action" value="sync_now">
                                    <input type="hidden" name="source_org_id" value="''' + str(source_org_id) + '''">
                                    <button type="submit" class="sync-btn sync-btn-success">
                                        <i class="fa fa-sync"></i> Perform Initial Sync
                                    </button>
                                </form>
                                <button onclick="history.back()" class="sync-btn sync-btn-default">
                                    <i class="fa fa-arrow-left"></i> Back to Menu
                                </button>
                            </div>
                        </div>
                    </div>
                    '''
                else:
                    raise Exception("Failed to save sync configuration")
                    
            except Exception as e:
                print get_common_css() + '''
                <div id="involvement-sync-container">
                    <div class="sync-alert sync-alert-danger">
                        <h4>Error Creating Sync Configuration</h4>
                        <p>''' + str(e) + '''</p>
                        <button onclick="history.back()" class="sync-btn sync-btn-default">
                            <i class="fa fa-arrow-left"></i> Back to Menu
                        </button>
                    </div>
                </div>
                '''
        
        elif action == "sync_all_manual":
            print show_loading_script()
            
            try:
                sync_enabled_orgs = sync_manager.get_sync_enabled_involvements()
                results_summary = []
                
                for org in sync_enabled_orgs:
                    settings = sync_manager.get_sync_settings(org.OrganizationId)
                    if settings['enabled'] and settings['target_org_id']:
                        try:
                            results = sync_manager.sync_members(org.OrganizationId, settings['target_org_id'])
                            source_name = sync_manager.get_involvement_name(org.OrganizationId)
                            target_name = sync_manager.get_involvement_name(settings['target_org_id'])
                            
                            results_summary.append({
                                'source': source_name,
                                'target': target_name,
                                'success': True,
                                'results': results
                            })
                        except Exception as e:
                            results_summary.append({
                                'source': org.OrganizationName,
                                'target': sync_manager.get_involvement_name(settings['target_org_id']) if settings['target_org_id'] else 'Unknown',
                                'success': False,
                                'error': str(e)
                            })
                
                print get_common_css() + '''
                <div id="involvement-sync-container">
                    <h3>Sync All Results</h3>
                '''
                
                for i, result in enumerate(results_summary):
                    panel_class = "sync-alert-success" if result['success'] else "sync-alert-danger"
                    print '''
                    <div class="sync-alert ''' + panel_class + '''">
                        <h4>''' + result['source'] + ''' → ''' + result['target'] + '''</h4>
                    '''
                    
                    if result['success']:
                        print '''
                        <div style="display: flex; gap: 15px; margin: 10px 0;">
                            <div style="text-align: center;">
                                <strong>''' + str(result['results']['members_added']) + '''</strong><br>
                                <small>Members Added</small>
                            </div>
                            <div style="text-align: center;">
                                <strong>''' + str(result['results']['members_updated']) + '''</strong><br>
                                <small>Members Updated</small>
                            </div>
                            <div style="text-align: center;">
                                <strong>''' + str(result['results']['members_removed']) + '''</strong><br>
                                <small>Members Removed</small>
                            </div>
                            <div style="text-align: center;">
                                <strong>''' + str(result['results']['subgroups_synced']) + '''</strong><br>
                                <small>Subgroups Added</small>
                            </div>
                            <div style="text-align: center;">
                                <strong>''' + str(result['results']['subgroups_removed']) + '''</strong><br>
                                <small>Subgroups Removed</small>
                            </div>
                            <div style="text-align: center;">
                                <strong>''' + str(len(result['results']['errors'])) + '''</strong><br>
                                <small>Errors</small>
                            </div>
                        </div>
                        '''
                        
                        if result['results']['errors']:
                            print '<div style="margin-top:10px;"><strong>Errors:</strong><ul>'
                            for error in result['results']['errors']:
                                print '<li>' + error + '</li>'
                            print '</ul></div>'
                    else:
                        print '<p><strong>Error:</strong> ' + result['error'] + '</p>'
                    
                    print '</div>'
                
                print '''
                    <div style="margin-top:20px;">
                        <button onclick="history.back()" class="sync-btn sync-btn-primary">
                            <i class="fa fa-arrow-left"></i> Back to Menu
                        </button>
                    </div>
                </div>
                '''
                
            except Exception as e:
                print get_common_css() + '''
                <div id="involvement-sync-container">
                    <div class="sync-alert sync-alert-danger">
                        <h4>Sync All Error</h4>
                        <p>''' + str(e) + '''</p>
                        <button onclick="history.back()" class="sync-btn sync-btn-primary">
                            <i class="fa fa-arrow-left"></i> Back to Menu
                        </button>
                    </div>
                </div>
                '''
        
        elif action == "sync_now" or action == "perform_manual_sync":
            print show_loading_script()
            
            try:
                source_org_id = int(str(model.Data.source_org_id))
                settings = sync_manager.get_sync_settings(source_org_id)
                
                if not settings['enabled']:
                    raise Exception("Sync is not enabled for this involvement (ID: " + str(source_org_id) + ")")
                
                if not settings['target_org_id']:
                    raise Exception("No target involvement configured for this involvement (ID: " + str(source_org_id) + ")")
                
                # Perform the sync
                results = sync_manager.sync_members(source_org_id, settings['target_org_id'])
                
                source_name = sync_manager.get_involvement_name(source_org_id)
                target_name = sync_manager.get_involvement_name(settings['target_org_id'])
                
                print get_common_css() + '''
                <div id="involvement-sync-container">
                    <h3>Sync Results: ''' + source_name + ''' → ''' + target_name + '''</h3>
                    ''' + sync_manager.render_sync_results(results) + '''
                    
                    <div style="margin-top:20px;">
                        <button onclick="history.back()" class="sync-btn sync-btn-primary">
                            <i class="fa fa-arrow-left"></i> Back to Menu
                        </button>
                    </div>
                </div>
                '''
                
            except Exception as e:
                print get_common_css() + '''
                <div id="involvement-sync-container">
                    <div class="sync-alert sync-alert-danger">
                        <h4>Sync Error</h4>
                        <p>''' + str(e) + '''</p>
                        <button onclick="history.back()" class="sync-btn sync-btn-primary">
                            <i class="fa fa-arrow-left"></i> Back to Menu
                        </button>
                    </div>
                </div>
                '''
        
        elif action == "edit_sync":
            print show_loading_script()
            source_org_id = int(str(model.Data.source_org_id))
            print render_setup_form(source_org_id)
            
        elif action == "update_sync":
            print show_loading_script()
            
            try:
                source_org_id = int(str(model.Data.source_org_id))
                target_org_id = int(str(model.Data.target_org_id))
                auto_sync = hasattr(model.Data, 'auto_sync') and str(model.Data.auto_sync) == "1"

                # Verify target org exists
                if not sync_manager.verify_organization_exists(target_org_id):
                    raise Exception("Target involvement ID " + str(target_org_id) + " does not exist in the database")

                if sync_manager.set_sync_settings(source_org_id, target_org_id, auto_sync):
                    source_name = sync_manager.get_involvement_name(source_org_id)
                    target_name = sync_manager.get_involvement_name(target_org_id)
                    
                    print get_common_css() + '''
                    <div id="involvement-sync-container">
                        <div class="sync-alert sync-alert-success">
                            <h4>Sync Configuration Updated Successfully!</h4>
                            <p><strong>Primary:</strong> ''' + source_name + '''</p>
                            <p><strong>Target:</strong> ''' + target_name + '''</p>
                            <p><strong>Auto Sync:</strong> ''' + ("Enabled" if auto_sync else "Disabled") + '''</p>
                            
                            <button onclick="history.back()" class="sync-btn sync-btn-primary">
                                <i class="fa fa-arrow-left"></i> Back to Menu
                            </button>
                        </div>
                    </div>
                    '''
                else:
                    raise Exception("Failed to update sync configuration")
                    
            except Exception as e:
                print get_common_css() + '''
                <div id="involvement-sync-container">
                    <div class="sync-alert sync-alert-danger">
                        <h4>Error Updating Sync Configuration</h4>
                        <p>''' + str(e) + '''</p>
                        <button onclick="history.back()" class="sync-btn sync-btn-primary">
                            <i class="fa fa-arrow-left"></i> Back to Menu
                        </button>
                    </div>
                </div>
                '''
        
        else:
            # Show main menu
            print show_loading_script()
            print render_main_menu()
    
    except Exception as e:
        # Print any errors for interactive use
        import traceback
        print "<h2>Error</h2>"
        print "<p>An error occurred: " + str(e) + "</p>"
        print "<pre>"
        traceback.print_exc()
        print "</pre>"
        print '<button onclick="history.back()" class="btn btn-primary">Back to Menu</button>'
