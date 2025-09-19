#roles=Edit

#####################################################################
# Registration Data Interactive Manager v2.9.5
#####################################################################
# Version 2.9.5 - Preload lookup values to avoid AJAX issues
# This tool provides an interactive interface for managing registration
# data without needing to export to CSV. Features include:
# 1. Interactive filtering and searching with column filters
# 2. Bulk selection of people
# 3. Tag assignment
# 4. Add to organization subgroups
# 5. Save question answers to extra values
# 6. Export with full question headers
# 7. Admin-only: Update standard person fields
#
# REQUIRED PERMISSION: Edit role to use the script
# Code visualization requires SuperAdmin role
# Person field updates require Admin role
#
# CHANGELOG v2.9:
# - Preload lookup values eliminating AJAX issues
# - Fixed AJAX handler JSON responses and parameter handling
# - Enhanced error handling with proper debugging
# - All lookup data embedded in page for instant dropdown population
#
# CHANGELOG v2.8:
# - Added Admin-only feature to update standard person fields
# - Support for updating Campus, Member Status, Position in Family
# - Fixed ag-Grid configuration and AJAX JSON handling
# - Dynamic field input types with automatic lookup value loading
#
# CHANGELOG v2.7.0:
# - Added ag-Grid Enterprise features including sidebar
# - Added column management panel to show/hide columns
# - Added row grouping capabilities for data analysis
# - Added Excel export functionality with full question headers
# - Added dedicated filters panel in sidebar
# - Fixed UserInRole to UserIsInRole API calls
# - Improved grid layout and performance
#
# CHANGELOG v2.6.0:
# - Added configuration class for easy customization
# - Implemented debug mode for administrators
# - Added critical database filters (IsDeceased, ArchivedFlag)
# - Added performance optimizations and pagination support
# - Improved error handling with standardized patterns
# - Added input sanitization functions
#
# CHANGELOG v2.5.1:
# - Changed main script permission to Edit role (via #roles directive)
# - Code visualization remains SuperAdmin only
# - Mermaid diagrams now collapsed by default (click to expand)
# - Removed redundant permission checks
#
# CHANGELOG v2.5:
# - Added Mermaid code structure visualization for SuperAdmin users
# - Visualization shows data flow and process architecture
# - Includes both graph and sequence diagram views
# - Script view permission restricted to SuperAdmin role
#
# CHANGELOG v2.4:
# - Organization dropdown now defaults to selected involvement if only one was queried
# - Enhanced feedback showing new members vs existing members when adding to org
# - Clarified that people are automatically added to org before subgroup assignment
# - Better status messages for organization operations
#
# CHANGELOG v2.3:
# - Fixed question column ordering to match reference table
# - Fixed subgroup creation using TouchPoint's AddSubGroup API
# - Subgroups now automatically created if they don't exist
# - Sorted question columns (Q1, Q2, Q3...) appear in numerical order
# - Export maintains correct question order
#
# CHANGELOG v2.2.1:
# - Fixed JavaScript regex escaping issues in cellRenderer
# - Corrected backslash handling in replace statements
# - Fixed "Uncaught SyntaxError" in value unescaping
#
# CHANGELOG v2.2:
# - Fixed JSON serialization issues with special characters
# - Fixed duplicate node ID errors for guest registrations
# - Fixed missing question data (Q1, Q2, etc.) not showing
# - Added better error handling for malformed data
# - Improved debug output for troubleshooting
#
# CHANGELOG v2.1:
# - Changed extra value saving to offer global value option instead of individual editing
# - Fixed subgroup functionality to work with both extra values and OrganizationMembers.SubGroup field
# - Subgroups now save as "OrgSubgroup_[OrgId]" for better clarity
# - Added preview of sample values when using original answers
# - Improved error messages and feedback
#
# CHANGELOG v2.0:
# - Fixed PyScript/PyScriptForm submission handling
# - Added per-column filtering in ag-Grid
# - Headers now show Q1, Q2, etc. with full questions in tooltips
# - Question reference table shows mapping
# - Export includes full question text as headers
#
#####################################################################
# Upload Instructions
#####################################################################
# To upload code to Touchpoint, use the following steps:
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File
# 3. Name the Python (suggested: RegDataManager) and paste all this code
# 4. Test and optionally add to menu
# 5. The script automatically requires Edit role (no need to set in menu)
#####################################################################

#written by: Ben Swaby 
#email: bswaby@fbchtn.or

# ===== CONFIGURATION SECTION =====
# Place all configuration at the top for easy customization and sharing
class Config:
    # Display settings
    PAGE_TITLE = "Registration Data Manager"
    GRID_HEIGHT = 600
    MAX_EXPORT_ROWS = 10000
    
    # Feature flags
    ENABLE_DEBUG = False  # Set to True for development
    ENABLE_MERMAID = True  # Show code visualization
    ENABLE_EXPORT = True
    ENABLE_BULK_OPERATIONS = True
    
    # Security settings
    # Note: Base Edit role is handled by #roles=Edit directive at top of script
    MERMAID_ROLE = "SuperAdmin"  # Role required for code visualization
    DEBUG_ROLE = "Admin"  # Role required for debug mode
    
    # Performance settings
    BATCH_SIZE = 100  # Batch size for bulk operations
    QUERY_TIMEOUT = 30  # seconds
    MAX_PEOPLE_PER_OPERATION = 1000  # Safety limit
    
    # Export settings
    EXPORT_FILENAME = "registration_data_export.csv"
    
    # Display preferences
    DATE_FORMAT = "MM/dd/yyyy"
    SHOW_GUEST_WARNINGS = True
    
    # Grid settings
    PAGE_SIZE = 200
    ENABLE_PAGINATION = True  # Enable for large datasets

# Handle AJAX requests FIRST (before any HTML output or model modifications)
# This handler processes lookup value requests for admin person field updates
try:
    import sys
    import os
    
    # Check if this is an AJAX request for lookup values
    if hasattr(model, 'Data') and hasattr(model.Data, 'action'):
        action = getattr(model.Data, 'action', '')
        if action == "get_lookup_values":
            # Get the field name (JavaScript sends 'field_name', not 'field')
            field_name = getattr(model.Data, 'field_name', '')
            
            # Build response
            result_values = []
            error_msg = ""
            debug_info = "Field requested: " + field_name
            
            try:
                # Define lookup queries with both schema prefixes
                rows = []
                query_used = ""
                
                if field_name == 'CampusId':
                    try:
                        query_used = "SELECT Id, Description FROM lookup.Campus ORDER BY Description"
                        rows = q.QuerySql(query_used)
                    except Exception as e1:
                        try:
                            query_used = "SELECT Id, Description FROM Campus ORDER BY Description"
                            rows = q.QuerySql(query_used)
                        except Exception as e2:
                            error_msg = "Campus query failed: " + str(e1) + " / " + str(e2)
                            
                elif field_name == 'MemberStatusId':
                    try:
                        query_used = "SELECT Id, Description FROM lookup.MemberStatus ORDER BY Id"
                        rows = q.QuerySql(query_used)
                    except Exception as e1:
                        try:
                            query_used = "SELECT Id, Description FROM MemberStatus ORDER BY Id"
                            rows = q.QuerySql(query_used)
                        except Exception as e2:
                            error_msg = "MemberStatus query failed: " + str(e1) + " / " + str(e2)
                            
                elif field_name == 'PositionInFamilyId':
                    try:
                        query_used = "SELECT Id, Description FROM lookup.FamilyPosition ORDER BY Description"
                        rows = q.QuerySql(query_used)
                    except Exception as e1:
                        try:
                            query_used = "SELECT Id, Description FROM FamilyPosition ORDER BY Description"
                            rows = q.QuerySql(query_used)
                        except Exception as e2:
                            error_msg = "FamilyPosition query failed: " + str(e1) + " / " + str(e2)
                            
                elif field_name == 'GenderId':
                    try:
                        query_used = "SELECT Id, Description FROM lookup.Gender ORDER BY Description"
                        rows = q.QuerySql(query_used)
                    except Exception as e1:
                        try:
                            query_used = "SELECT Id, Description FROM Gender ORDER BY Description"
                            rows = q.QuerySql(query_used)
                        except Exception as e2:
                            error_msg = "Gender query failed: " + str(e1) + " / " + str(e2)
                            
                elif field_name == 'MaritalStatusId':
                    try:
                        query_used = "SELECT Id, Description FROM lookup.MaritalStatus ORDER BY Id"
                        rows = q.QuerySql(query_used)
                    except Exception as e1:
                        try:
                            query_used = "SELECT Id, Description FROM MaritalStatus ORDER BY Id"
                            rows = q.QuerySql(query_used)
                        except Exception as e2:
                            error_msg = "MaritalStatus query failed: " + str(e1) + " / " + str(e2)
                else:
                    rows = []
                    debug_info += " (field not recognized: " + field_name + ")"
                
                # Add query info to debug
                if query_used:
                    debug_info += " Query: " + query_used
                
                # Process results
                if rows:
                    debug_info += " Rows returned: " + str(len(rows))
                    for row in rows:
                        # Try different attribute name combinations
                        id_value = None
                        desc_value = None
                        
                        # Check for Id/ID variations
                        if hasattr(row, 'Id'):
                            id_value = str(row.Id)
                        elif hasattr(row, 'ID'):
                            id_value = str(row.ID)
                        elif hasattr(row, 'id'):
                            id_value = str(row.id)
                            
                        # Check for Description variations
                        if hasattr(row, 'Description'):
                            desc_value = str(row.Description)
                        elif hasattr(row, 'description'):
                            desc_value = str(row.description)
                        elif hasattr(row, 'Name'):
                            desc_value = str(row.Name)
                        elif hasattr(row, 'name'):
                            desc_value = str(row.name)
                            
                        if id_value and desc_value:
                            result_values.append({
                                "id": id_value,
                                "name": desc_value
                            })
                        else:
                            # Debug what attributes the row has
                            attrs = []
                            for attr in dir(row):
                                if not attr.startswith('_'):
                                    attrs.append(attr + "=" + str(getattr(row, attr, 'N/A'))[:20])
                            debug_info += " Row attrs: " + ", ".join(attrs)
                
                debug_info += " - Found " + str(len(result_values)) + " values"
                
            except Exception as e:
                error_msg = str(e).replace('"', '\\"').replace('\n', ' ')
                debug_info += " - Error: " + error_msg
            
            # Output JSON response manually
            print("{")
            print('  "values": [')
            for i, val in enumerate(result_values):
                print('    {"id": "' + val["id"] + '", "name": "' + val["name"].replace('"', '\\"') + '"}')
                if i < len(result_values) - 1:
                    print(",")
            print("  ],")
            print('  "debug": "' + debug_info.replace('"', '\\"') + '"')
            if error_msg:
                print(',')
                print('  "error": "' + error_msg + '"')
            print("}")
            
            # Exit immediately
            sys.stdout.flush()
            os._exit(0)
            
except Exception as outer_error:
    # If ANY error occurs in the AJAX handler, output valid JSON and exit
    import sys
    import os
    print('{"values": [], "error": "Handler exception", "debug": "' + str(outer_error).replace('"', '\\"') + '"}')
    sys.stdout.flush()
    os._exit(0)

# Set page header from config (only after AJAX check)
model.Header = Config.PAGE_TITLE

# ===== HELPER FUNCTIONS =====

def debug_print(message, data=None):
    """Print debug information - Admin only"""
    if not Config.ENABLE_DEBUG or not model.UserIsInRole(Config.DEBUG_ROLE):
        return
    
    print("""
    <div class="debug-info" style="background: #fffbf0; border: 1px solid #ffa500; 
         padding: 10px; margin: 10px 0; font-family: monospace; font-size: 12px;">
        <strong>DEBUG:</strong> {0}
    """.format(message))
    
    if data:
        import json
        try:
            print("<pre>{0}</pre>".format(json.dumps(data, indent=2)))
        except:
            print("<pre>{0}</pre>".format(str(data)))
    
    print("</div>")

def print_error(section, error, show_traceback=True):
    """Standardized error display with optional traceback"""
    print("""
    <div class="alert alert-danger">
        <h4><i class="glyphicon glyphicon-exclamation-sign"></i> Error in {0}</h4>
        <p>{1}</p>
    """.format(section, str(error)))
    
    if show_traceback and (Config.ENABLE_DEBUG or model.UserIsInRole(Config.DEBUG_ROLE)):
        import traceback
        print("<pre style='max-height: 200px; overflow-y: auto; font-size: 11px;'>{0}</pre>".format(
            traceback.format_exc()
        ))
    
    print("</div>")

def sanitize_sql_input(value):
    """Sanitize input for SQL queries"""
    if value is None:
        return ''
    
    # Convert to string and clean
    value = str(value)
    
    # Remove/escape dangerous characters
    value = value.replace("'", "''")  # Escape single quotes
    value = value.replace(";", "")     # Remove semicolons
    value = value.replace("--", "")    # Remove SQL comments
    value = value.replace("/*", "")    # Remove block comments
    value = value.replace("*/", "")
    value = value.replace("xp_", "")   # Remove extended procedures
    value = value.replace("sp_", "")   # Remove stored procedures
    
    return value

def show_debug_panel():
    """Show debug information panel - Admin only"""
    if not model.UserIsInRole(Config.DEBUG_ROLE):
        return
    
    import json
    
    # Get request data
    request_data = {}
    if hasattr(model, 'Data'):
        for attr in dir(model.Data):
            if not attr.startswith('_'):
                try:
                    request_data[attr] = getattr(model.Data, attr)
                except:
                    request_data[attr] = 'Unable to retrieve'
    
    print("""
    <div class="panel panel-warning" id="debugPanel" style="margin-top: 20px;">
        <div class="panel-heading">
            <h3 class="panel-title">
                <i class="fa fa-bug"></i> Debug Information 
                <small>(Admin Only)</small>
                <button class="btn btn-xs btn-warning pull-right" 
                        onclick="$('#debugContent').toggle()">
                    Toggle
                </button>
            </h3>
        </div>
        <div class="panel-body" id="debugContent" style="display: none;">
            <h4>Configuration</h4>
            <pre>{0}</pre>
            
            <h4>User Context</h4>
            <pre>UserName: {1}
UserPeopleId: {2}
Has Edit Role: {3}
UserIsInRole(Admin): {4}
UserIsInRole(SuperAdmin): {5}</pre>
            
            <h4>Request Data</h4>
            <pre>{6}</pre>
        </div>
    </div>
    """.format(
        json.dumps({
            'Debug Mode': Config.ENABLE_DEBUG,
            'Batch Size': Config.BATCH_SIZE,
            'Max Export Rows': Config.MAX_EXPORT_ROWS,
            'Pagination Enabled': Config.ENABLE_PAGINATION
        }, indent=2),
        model.UserName,
        model.UserPeopleId,
        "Yes (handled by #roles directive)",
        model.UserIsInRole("Admin"),
        model.UserIsInRole("SuperAdmin"),
        json.dumps(request_data, indent=2)
    ))

# Original helper functions (kept for backward compatibility)

def get_org_name(org_id):
    try:
        sql = "SELECT OrganizationName FROM Organizations WHERE OrganizationId = " + str(org_id)
        result = q.QuerySqlTop1(sql)
        if result and hasattr(result, 'OrganizationName'):
            return result.OrganizationName
        return "Unknown Organization"
    except:
        return "Organization #" + str(org_id)

def get_registration_orgs():
    try:
        # First try with program information
        sql = """
        SELECT DISTINCT 
            o.OrganizationId, 
            o.OrganizationName,
            ISNULL(p.Name, 'No Program') as ProgramName,
            ISNULL(p.Id, 0) as ProgramId
        FROM Organizations o
        JOIN RegQuestion rq ON o.OrganizationId = rq.OrganizationId
        LEFT JOIN Division d ON o.DivisionId = d.Id
        LEFT JOIN Program p ON d.ProgId = p.Id
        WHERE o.OrganizationStatusId = 30 
        ORDER BY ISNULL(p.Name, 'No Program'), o.OrganizationName
        """
        results = q.QuerySql(sql)
        debug_print("Registration orgs query with programs succeeded", {"count": len(results) if results else 0})
        return results
    except Exception as e:
        debug_print("Program query failed, trying simple query", {"error": str(e)})
        # Fallback to simple query without program information
        try:
            sql = """
            SELECT DISTINCT o.OrganizationId, o.OrganizationName
            FROM Organizations o
            JOIN RegQuestion rq ON o.OrganizationId = rq.OrganizationId
            WHERE o.OrganizationStatusId = 30 
            ORDER BY o.OrganizationName
            """
            results = q.QuerySql(sql)
            # Add empty program info for compatibility
            if results:
                for org in results:
                    org.ProgramName = 'No Program'
                    org.ProgramId = 0
            debug_print("Simple registration orgs query succeeded", {"count": len(results) if results else 0})
            return results
        except Exception as e2:
            debug_print("Both queries failed", {"error1": str(e), "error2": str(e2)})
            return []

# ===== MAIN EXECUTION =====
def main():
    # Get organization IDs from parameter (p1) or set defaults
    org_ids = model.Data.p1 if hasattr(model.Data, 'p1') and model.Data.p1 else ""
    
    # Handle form submissions for bulk actions
    action = ""
    selected_people = ""
    if hasattr(model.Data, 'action'):
        action = str(model.Data.action) if model.Data.action else ""
        selected_people = str(model.Data.selected_people) if hasattr(model.Data, 'selected_people') else ""

    debug_print("Form submission received", {
        "action": action,
        "selected_people_count": len(selected_people.split(',')) if selected_people else 0
    })

    if action and selected_people:
        people_ids = [int(pid.strip()) for pid in selected_people.split(',') if pid.strip() and pid.strip().isdigit()]
        
        # Safety check for maximum people per operation
        if len(people_ids) > Config.MAX_PEOPLE_PER_OPERATION:
            print('<div class="alert alert-warning">Operation limited to {0} people at a time. Please select fewer people.</div>'.format(
                Config.MAX_PEOPLE_PER_OPERATION))
            people_ids = people_ids[:Config.MAX_PEOPLE_PER_OPERATION]
        
        debug_print("Processing bulk action", {
            "action": action,
            "people_count": len(people_ids),
            "sample_ids": people_ids[:5]
        })
        
        if action == "tag":
            tag_name = str(model.Data.tag_name) if hasattr(model.Data, 'tag_name') else ""
            if tag_name and people_ids:
                # Create or get tag
                tag_id = model.CreateQueryTag(tag_name, "peopleids='{}'".format(','.join(map(str, people_ids))))
                print('<div class="alert alert-success">Successfully tagged {} people with "{}"</div>'.format(len(people_ids), tag_name))
        
        elif action == "add_to_org":
            target_org_id = str(model.Data.target_org_id) if hasattr(model.Data, 'target_org_id') else ""
            subgroup = str(model.Data.subgroup) if hasattr(model.Data, 'subgroup') else ""
            if target_org_id and people_ids:
                success_count = 0
                already_member_count = 0
                skipped_guests = 0
                for person_id in people_ids:
                    if person_id <= 0:  # Skip guest records
                        skipped_guests += 1
                        continue
                    try:
                        # Check if already a member
                        was_member = model.InOrg(person_id, int(target_org_id))
                        
                        # Add member to organization (does nothing if already member)
                        model.AddMemberToOrg(person_id, int(target_org_id))
                        
                        if was_member:
                            already_member_count += 1
                        else:
                            success_count += 1
                        
                        # If subgroup is specified, add to subgroup
                        if subgroup:
                            # Use TouchPoint's AddSubGroup function which creates the subgroup if needed
                            model.AddSubGroup(person_id, int(target_org_id), subgroup)
                            
                            # Also save as extra value for easy reference
                            model.AddExtraValueText(person_id, "OrgSubgroup_{}".format(target_org_id), subgroup)
                    except Exception as e:
                        print('<div class="alert alert-warning">Could not add person {}: {}</div>'.format(person_id, str(e)))
                        
                if success_count > 0 or already_member_count > 0:
                    if subgroup:
                        print('<div class="alert alert-success">')
                        if success_count > 0:
                            print('Added {} new people to organization<br>'.format(success_count))
                        if already_member_count > 0:
                            print('Updated {} existing members<br>'.format(already_member_count))
                        print('All assigned to subgroup "{}"<br>'.format(subgroup))
                        print('Subgroup tracked in TouchPoint and as extra value: OrgSubgroup_{}'.format(target_org_id))
                        print('</div>')
                    else:
                        print('<div class="alert alert-success">')
                        if success_count > 0:
                            print('Added {} new people to organization<br>'.format(success_count))
                        if already_member_count > 0:
                            print('{} were already members (no change)<br>'.format(already_member_count))
                        print('</div>')
                
                if skipped_guests > 0:
                    print('<div class="alert alert-info">Skipped {} guest registrations (they must be converted to people records first)</div>'.format(skipped_guests))
        
        elif action == "save_to_extravalue":
            question_col = str(model.Data.question_col) if hasattr(model.Data, 'question_col') else ""
            extravalue_field = str(model.Data.extravalue_field) if hasattr(model.Data, 'extravalue_field') else ""
            answers_json = str(model.Data.answers_json) if hasattr(model.Data, 'answers_json') else "{}"
            use_global = str(model.Data.use_global_value) if hasattr(model.Data, 'use_global_value') else ""
            global_value = str(model.Data.global_value) if hasattr(model.Data, 'global_value') else ""
            
            if question_col and extravalue_field:
                import json
                try:
                    # Parse the JSON data
                    answers = json.loads(answers_json)
                    success_count = 0
                    error_count = 0
                    skipped_guests = 0
                    
                    for person_id in people_ids:
                        if person_id <= 0:  # Skip guest records
                            skipped_guests += 1
                            continue
                        try:
                            person_id_str = str(person_id)
                            
                            # Use global value if specified, otherwise use original answer
                            if use_global == "1" and global_value:
                                answer_value = global_value
                            elif person_id_str in answers:
                                answer_value = answers[person_id_str]
                            else:
                                answer_value = None
                                
                            if answer_value:
                                # Clean the answer value before saving
                                answer_value = answer_value.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
                                model.AddExtraValueText(person_id, extravalue_field, answer_value)
                                success_count += 1
                        except Exception as person_error:
                            error_count += 1
                            print('<div class="alert alert-warning">Warning: Could not save value for person ID {}: {}</div>'.format(person_id, str(person_error)))
                    
                    if success_count > 0:
                        if use_global == "1" and global_value:
                            print('<div class="alert alert-success">Successfully saved global value "{}" to {} people in extra value field "{}"</div>'.format(
                                global_value, success_count, extravalue_field))
                        else:
                            print('<div class="alert alert-success">Successfully saved {} values to extra value field "{}"</div>'.format(
                                success_count, extravalue_field))
                    
                    if error_count > 0:
                        print('<div class="alert alert-warning">{} values could not be saved due to errors</div>'.format(error_count))
                    
                    if skipped_guests > 0:
                        print('<div class="alert alert-info">Skipped {} guest registrations (they must be converted to people records first)</div>'.format(skipped_guests))
                        
                except Exception as e:
                    print('<div class="alert alert-danger">Error processing extra value save: {}</div>'.format(str(e)))
        
        elif action == "update_person_field":
            # Admin-only: Update person fields
            if model.UserIsInRole("Admin"):
                field_name = str(model.Data.field_name) if hasattr(model.Data, 'field_name') else ""
                field_value = str(model.Data.field_value_final) if hasattr(model.Data, 'field_value_final') else ""
                use_current_value = str(model.Data.use_current_value) if hasattr(model.Data, 'use_current_value') else "0"
                
                if field_name and people_ids:
                    success_count = 0
                    error_count = 0
                    skipped_guests = 0
                    
                    for person_id in people_ids:
                        if person_id <= 0:  # Skip guest records
                            skipped_guests += 1
                            continue
                        
                        try:
                            # Get the person object
                            person = model.GetPerson(person_id)
                            if person:
                                # If using current value, keep existing value (essentially a no-op)
                                if use_current_value == "1":
                                    # Just touching the record without changing the field
                                    # This allows for batch operations where you might want to preserve current values
                                    success_count += 1
                                else:
                                    # Update the field based on type
                                    if field_name in ['CampusId', 'MemberStatusId', 'PositionInFamilyId', 'GenderId', 'MaritalStatusId']:
                                        # ID fields - convert to integer
                                        setattr(person, field_name, int(field_value) if field_value else 0)
                                    elif field_name in ['DoNotCallFlag', 'DoNotMailFlag', 'DoNotVisitFlag', 'DoNotPublishPhones']:
                                        # Boolean fields
                                        setattr(person, field_name, field_value == '1')
                                    else:
                                        # Text fields
                                        setattr(person, field_name, field_value)
                                    
                                    # Save the person record
                                    model.UpdatePerson(person)
                                    success_count += 1
                            else:
                                error_count += 1
                                
                        except Exception as person_error:
                            error_count += 1
                            debug_print("Error updating person {}".format(person_id), {"error": str(person_error)})
                    
                    # Display results
                    if success_count > 0:
                        if use_current_value == "1":
                            print('<div class="alert alert-success">Successfully preserved current {} values for {} people</div>'.format(
                                field_name, success_count))
                        else:
                            print('<div class="alert alert-success">Successfully updated {} field for {} people</div>'.format(
                                field_name, success_count))
                    
                    if error_count > 0:
                        print('<div class="alert alert-warning">{} people could not be updated due to errors</div>'.format(error_count))
                    
                    if skipped_guests > 0:
                        print('<div class="alert alert-info">Skipped {} guest registrations (they must be converted to people records first)</div>'.format(skipped_guests))
                else:
                    print('<div class="alert alert-danger">Missing required field name or people selection</div>')
            else:
                print('<div class="alert alert-danger">This action requires Administrator permissions</div>')
    
    # Display organization selection form
    if not org_ids:
        print("""
        <h2>Select Involvements for Registration Data Manager</h2>
        <form method="get" action="">
            <div style="margin-bottom: 20px;">
                <label for="org_ids">Enter Involvement IDs (comma-separated):</label>
                <input type="text" id="org_ids" name="p1" style="width: 300px;">
            </div>
            <p>- OR -</p>
            <div style="margin-bottom: 20px;">
                <label>Select from involvements with registration questions:</label>
                <div style="max-height: 300px; overflow-y: auto; border: 1px solid #ccc; padding: 10px; margin-top: 5px;">
        """)
        
        orgs = get_registration_orgs()
        # Build debug data safely
        debug_data = {"count": len(orgs) if orgs else 0}
        if orgs:
            # Get first 5 orgs for debugging
            sample_orgs = []
            for i, org in enumerate(orgs):
                if i >= 5:
                    break
                sample_orgs.append({"id": org.OrganizationId, "name": org.OrganizationName})
            debug_data["orgs"] = sample_orgs
        else:
            debug_data["orgs"] = []
        
        debug_print("Registration orgs for selection", debug_data)
        if orgs:
            # Group organizations by program
            current_program = None
            for org in orgs:
                # Check if we're starting a new program group
                if current_program != org.ProgramName:
                    if current_program is not None:
                        print("</div>")  # Close previous program group
                    current_program = org.ProgramName
                    print("""
                    <div style="margin-bottom: 15px;">
                        <h4 style="margin: 5px 0; color: #666;">{0}</h4>
                    """.format(current_program))
                
                print("""
                <div style="margin-bottom: 5px; margin-left: 15px;">
                    <a href="?p1={0}" style="text-decoration: none;">
                        <button type="button" style="width: calc(100% - 15px); text-align: left; padding: 5px;">
                            {1} (ID: {0})
                        </button>
                    </a>
                </div>
                """.format(org.OrganizationId, org.OrganizationName))
            
            if current_program is not None:
                print("</div>")  # Close last program group
        else:
            print("<p>No involvements with registration questions found.</p>")
        
        print("""
                </div>
            </div>
            <div>
                <button type="submit">Launch Manager</button>
            </div>
        </form>
        """)
    
    # Process data when org_ids are provided
    else:
        try:
            # Get available organizations with program information for adding people
            try:
                # First try with program information
                all_orgs = q.QuerySql("""
                    SELECT 
                        o.OrganizationId, 
                        o.OrganizationName,
                        ISNULL(p.Name, 'No Program') as ProgramName,
                        ISNULL(p.Id, 0) as ProgramId
                    FROM Organizations o
                    LEFT JOIN Division d ON o.DivisionId = d.Id
                    LEFT JOIN Program p ON d.ProgId = p.Id
                    WHERE o.OrganizationStatusId = 30 
                    ORDER BY ISNULL(p.Name, 'No Program'), o.OrganizationName
                """)
                debug_print("All orgs query with programs succeeded", {"count": len(all_orgs) if all_orgs else 0})
            except Exception as e:
                debug_print("Program query failed for all orgs, trying simple query", {"error": str(e)})
                # Fallback to simple query
                all_orgs = q.QuerySql("""
                    SELECT 
                        o.OrganizationId, 
                        o.OrganizationName
                    FROM Organizations o
                    WHERE o.OrganizationStatusId = 30 
                    ORDER BY o.OrganizationName
                """)
                # Add empty program info for compatibility
                if all_orgs:
                    for org in all_orgs:
                        org.ProgramName = 'No Program'
                        org.ProgramId = 0
        
            # Display header
            print("""<h2>Registration Data Manager</h2>""")
            print('<h3>Selected Involvements:</h3>')
            print('<ul>')
        
            org_id_list = org_ids.split(",")
            for org_id in org_id_list:
                if org_id and org_id.isdigit():
                    org_name = get_org_name(org_id)
                    print('<li>{0} (ID: {1})</li>'.format(org_name, org_id))
            print('</ul>')
        
            # Get questions consolidated by label
            question_sql = """
            WITH QuestionGroups AS (
                SELECT 
                    Label,
                    MIN([Order]) AS DefaultOrder,
                    COUNT(DISTINCT OrganizationId) AS OrgCount
                FROM RegQuestion
                WHERE OrganizationId IN ({0})
                GROUP BY Label
            )
            SELECT 
                ROW_NUMBER() OVER (ORDER BY DefaultOrder) AS QuestionID,
                Label,
                DefaultOrder,
                OrgCount
            FROM QuestionGroups
            ORDER BY DefaultOrder;
            """.format(org_ids)
        
            questions = q.QuerySql(question_sql)
        
            if not questions:
                show_error("No registration questions found for the selected involvements.")
                print('<p><a href="?" class="btn btn-primary">Select Different Organizations</a></p>')
            else:
                # Define mapping for SQL column names
                question_columns = []
                question_headers = []
                column_map = {}
            
                # Debug: Print questions found
                print("<!-- Debug: Found {} questions -->".format(len(questions) if questions else 0))
            
                for q_idx, q_item in enumerate(questions):
                    column_name = "Q{0}".format(q_idx + 1)
                    question_columns.append(column_name)
                    question_headers.append(q_item.Label)
                    column_map[q_item.Label] = column_name
                    print("<!-- Debug: Question {}: {} -> {} -->".format(q_idx + 1, q_item.Label[:50], column_name))
            
                # Get registration data with critical filters
                data_sql = """
                SELECT 
                    r.OrganizationId,
                    o.OrganizationName,
                    COALESCE(rp.PeopleId, 0) AS PeopleId,
                    rp.RegPeopleId,
                    COALESCE(p.FirstName, rp.FirstName) AS FirstName,
                    COALESCE(p.LastName, rp.LastName) AS LastName,
                    COALESCE(p.EmailAddress, rp.Email) AS EmailAddress,
                    COALESCE(p.HomePhone, rp.Phone) AS Phone,
                    rq.Label AS QuestionLabel,
                    ra.AnswerValue
                FROM Registration r
                JOIN RegPeople rp ON r.RegistrationId = rp.RegistrationId
                JOIN Organizations o ON r.OrganizationId = o.OrganizationId
                LEFT JOIN People p ON rp.PeopleId = p.PeopleId 
                    AND p.IsDeceased = 0 
                    AND p.ArchivedFlag = 0
                LEFT JOIN RegAnswer ra ON rp.RegPeopleId = ra.RegPeopleId
                LEFT JOIN RegQuestion rq ON ra.RegQuestionId = rq.RegQuestionId
                WHERE r.OrganizationId IN ({0})
                """.format(org_ids)
            
                debug_print("Executing registration data query", {
                    "org_ids": org_ids,
                    "filters": "IsDeceased=0, ArchivedFlag=0"
                })
            
                reg_data = q.QuerySql(data_sql)
            
                # Process data into person-centric dictionary
                person_data = {}
                guest_counter = 1
            
                print("<!-- Debug: Processing {} registration records -->".format(len(reg_data) if reg_data else 0))
            
                if reg_data:
                    for record in reg_data:
                        # Create unique key for each person/guest
                        if record.PeopleId and record.PeopleId > 0:
                            person_key = str(record.PeopleId)
                            unique_id = record.PeopleId
                        else:
                            # Use negative IDs for guests to ensure uniqueness
                            person_key = "guest_" + str(record.RegPeopleId)
                            unique_id = -guest_counter
                            guest_counter += 1
                    
                        if person_key not in person_data:
                            person_data[person_key] = {
                                'OrganizationId': record.OrganizationId,
                                'OrganizationName': record.OrganizationName or '',
                                'PeopleId': unique_id,  # Use unique ID instead of 0
                                'FirstName': record.FirstName or '',
                                'LastName': record.LastName or '',
                                'EmailAddress': record.EmailAddress or '',
                                'Phone': record.Phone or ''
                            }
                            # Initialize all question columns
                            for col in question_columns:
                                person_data[person_key][col] = ''
                    
                        # Map question to column - only if we have a question label
                        if record.QuestionLabel and record.QuestionLabel in column_map:
                            column = column_map[record.QuestionLabel]
                            answer_value = record.AnswerValue
                            if answer_value is not None:
                                # Clean answer value to prevent JSON issues
                                answer_str = str(answer_value)
                                # Escape special characters that could break JSON
                                answer_str = answer_str.replace('\\', '\\\\')  # Escape backslashes first
                                answer_str = answer_str.replace('"', '\\"')    # Escape quotes
                                answer_str = answer_str.replace('\n', '\\n')   # Escape newlines
                                answer_str = answer_str.replace('\r', '\\r')   # Escape carriage returns
                                answer_str = answer_str.replace('\t', '\\t')   # Escape tabs
                                person_data[person_key][column] = answer_str
                                print("<!-- Debug: Set {} for person {} to: {} -->".format(column, person_key, answer_str[:50]))
            
                print("<!-- Debug: Processed {} unique people/guests -->".format(len(person_data)))
            
                # Convert to list for JSON serialization
                grid_data = []
                for person_key, data in person_data.items():
                    grid_data.append(data)
            
                # Import json for proper serialization
                import json
            
                # Create question mapping for export
                question_mapping = {}
                for idx, label in enumerate(question_headers):
                    question_mapping["Q{}".format(idx + 1)] = label
            
                # Pre-load all lookup values for JavaScript (Admin feature)
                lookup_values = {}
                if model.UserIsInRole("Admin"):
                    try:
                        # Campus
                        campus_data = []
                        try:
                            campus_rows = q.QuerySql("SELECT Id, Description FROM lookup.Campus ORDER BY Description")
                            for row in campus_rows:
                                campus_data.append({"id": str(row.Id), "name": str(row.Description)})
                        except:
                            try:
                                campus_rows = q.QuerySql("SELECT Id, Description FROM Campus ORDER BY Description")
                                for row in campus_rows:
                                    campus_data.append({"id": str(row.Id), "name": str(row.Description)})
                            except:
                                pass
                        lookup_values["CampusId"] = campus_data
                        
                        # Member Status
                        member_status_data = []
                        try:
                            ms_rows = q.QuerySql("SELECT Id, Description FROM lookup.MemberStatus ORDER BY Id")
                            for row in ms_rows:
                                member_status_data.append({"id": str(row.Id), "name": str(row.Description)})
                        except:
                            try:
                                ms_rows = q.QuerySql("SELECT Id, Description FROM MemberStatus ORDER BY Id")
                                for row in ms_rows:
                                    member_status_data.append({"id": str(row.Id), "name": str(row.Description)})
                            except:
                                pass
                        lookup_values["MemberStatusId"] = member_status_data
                        
                        # Position in Family
                        position_data = []
                        try:
                            pos_rows = q.QuerySql("SELECT Id, Description FROM lookup.FamilyPosition ORDER BY Description")
                            for row in pos_rows:
                                position_data.append({"id": str(row.Id), "name": str(row.Description)})
                        except:
                            try:
                                pos_rows = q.QuerySql("SELECT Id, Description FROM FamilyPosition ORDER BY Description")
                                for row in pos_rows:
                                    position_data.append({"id": str(row.Id), "name": str(row.Description)})
                            except:
                                pass
                        lookup_values["PositionInFamilyId"] = position_data
                        
                        # Gender
                        gender_data = []
                        try:
                            gender_rows = q.QuerySql("SELECT Id, Description FROM lookup.Gender ORDER BY Description")
                            for row in gender_rows:
                                gender_data.append({"id": str(row.Id), "name": str(row.Description)})
                        except:
                            try:
                                gender_rows = q.QuerySql("SELECT Id, Description FROM Gender ORDER BY Description")
                                for row in gender_rows:
                                    gender_data.append({"id": str(row.Id), "name": str(row.Description)})
                            except:
                                pass
                        lookup_values["GenderId"] = gender_data
                        
                        # Marital Status
                        marital_data = []
                        try:
                            marital_rows = q.QuerySql("SELECT Id, Description FROM lookup.MaritalStatus ORDER BY Id")
                            for row in marital_rows:
                                marital_data.append({"id": str(row.Id), "name": str(row.Description)})
                        except:
                            try:
                                marital_rows = q.QuerySql("SELECT Id, Description FROM MaritalStatus ORDER BY Id")
                                for row in marital_rows:
                                    marital_data.append({"id": str(row.Id), "name": str(row.Description)})
                            except:
                                pass
                        lookup_values["MaritalStatusId"] = marital_data
                        
                    except Exception as lookup_error:
                        debug_print("Error loading lookup values", {"error": str(lookup_error)})
            
                # Generate interactive interface
                print("""
                <!-- Registration Data Manager v2.9.5 -->
                <!-- Load required libraries - using Enterprise version for sidebar and Excel export -->
                <link rel="stylesheet" href="https://unpkg.com/ag-grid-enterprise@25.3.0/dist/styles/ag-grid.css">
                <link rel="stylesheet" href="https://unpkg.com/ag-grid-enterprise@25.3.0/dist/styles/ag-theme-alpine.css">
                <script src="https://unpkg.com/ag-grid-enterprise@25.3.0/dist/ag-grid-enterprise.min.noStyle.js"></script>
            
                <style>
                .manager-container {
                    margin: 20px 0;
                    overflow-x: hidden;
                }
                /* Grid wrapper to prevent overflow */
                .grid-wrapper {
                    position: relative;
                    width: 100%;
                    clear: both;
                }
                .action-panel {
                    background: #f5f5f5;
                    padding: 15px;
                    border-radius: 5px;
                    margin-bottom: 20px;
                }
                .action-section {
                    margin-bottom: 15px;
                    padding: 10px;
                    background: white;
                    border-radius: 3px;
                }
                .action-section h4 {
                    margin-top: 0;
                    color: #333;
                }
                .filter-panel {
                    background: #e8f0fe;
                    padding: 15px;
                    border-radius: 5px;
                    margin-bottom: 20px;
                }
                .filter-input {
                    width: 200px;
                    padding: 5px;
                    margin-right: 10px;
                }
                #myGrid {
                    height: """ + str(Config.GRID_HEIGHT) + """px;
                    width: 100%;
                    position: relative;
                    overflow: hidden;
                    margin-bottom: 20px;
                }
                /* Grid container styling for sidebar */
                .ag-root-wrapper {
                    border: 1px solid #ddd;
                    height: 100%;
                }
                /* Ensure sidebar appears properly */
                .ag-side-bar {
                    border-left: 1px solid #ddd;
                }
                .ag-side-bar .ag-side-button {
                    border-top: 1px solid #ddd;
                }
                /* Ensure the grid viewport doesn't overflow */
                .ag-body-viewport {
                    overflow-y: auto !important;
                }
                .selection-count {
                    font-weight: bold;
                    color: #1976d2;
                    font-size: 18px;
                }
                .btn {
                    padding: 8px 16px;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    margin-right: 10px;
                }
                .btn-primary { background: #1976d2; color: white; }
                .btn-success { background: #4caf50; color: white; }
                .btn-warning { background: #ff9800; color: white; }
                .btn-info { background: #2196f3; color: white; }
                .btn:hover { opacity: 0.9; }
            
                /* Question reference table */
                .question-reference {
                    background: #f9f9f9;
                    padding: 15px;
                    margin-bottom: 20px;
                    border: 1px solid #ddd;
                    border-radius: 5px;
                }
                .question-reference table {
                    width: 100%;
                    border-collapse: collapse;
                }
                .question-reference th {
                    background: #f0f0f0;
                    padding: 8px;
                    text-align: left;
                }
                .question-reference td {
                    padding: 8px;
                    border-top: 1px solid #ddd;
                }
                .guest-indicator {
                    color: #666;
                    font-style: italic;
                    font-size: 11px;
                }
                /* Row grouping panel styling */
                .ag-row-group-panel {
                    background: #f5f5f5;
                    border: 1px solid #ddd;
                    padding: 10px;
                    margin-bottom: 10px;
                    border-radius: 4px;
                }
                .ag-row-group-panel::before {
                    content: "Drag columns here to create row groups:";
                    display: block;
                    margin-bottom: 5px;
                    font-weight: bold;
                    color: #666;
                }
                .help-text {
                    display: inline-block;
                    margin-top: 5px;
                }
                /* Sidebar styling */
                .ag-tool-panel-wrapper {
                    border-left: 1px solid #ddd;
                    background: #f9f9f9;
                }
                </style>
            
                <div class="manager-container">
                    <!-- Question Reference -->
                    <div class="question-reference">
                        <h4>Question Reference</h4>
                        <table>
                            <tr>
                                <th>Column</th>
                                <th>Question</th>
                                <th>Used in # Orgs</th>
                            </tr>
                """)
            
                for q_idx, q_item in enumerate(questions):
                    column_name = "Q{0}".format(q_idx + 1)
                    print('<tr><td>{0}</td><td>{1}</td><td style="text-align: center;">{2}</td></tr>'.format(
                        column_name, q_item.Label, q_item.OrgCount))
            
                print("""
                        </table>
                    </div>
                
                    <!-- Filter Panel -->
                    <div class="filter-panel">
                        <h4>Filters & Controls</h4>
                        <input type="text" id="quickFilter" class="filter-input" placeholder="Search all columns..." onkeyup="onQuickFilter()">
                        <button onclick="clearFilters()" class="btn btn-info">Clear All Filters</button>
                        <button onclick="exportSelected()" class="btn btn-success">Export Selected (CSV)</button>
                        <button onclick="exportToExcel()" class="btn btn-success" title="Export to Excel">
                            <i class="fa fa-file-excel-o"></i> Export to Excel
                        </button>
                        <button onclick="toggleSidebar()" class="btn btn-warning" title="Show/hide column controls and grouping">
                            <i class="fa fa-columns"></i> Column Manager
                        </button>
                        <span class="help-text" style="margin-left: 10px; font-size: 12px; color: #666;">
                            <i class="fa fa-info-circle"></i> Use Column Manager to show/hide columns, create groups, and apply filters
                        </span>
                    </div>
                
                    <!-- Action Panel -->
                    <div class="action-panel">
                        <h3>Bulk Actions</h3>
                        <p>Selected: <span id="selectedCount" class="selection-count">0</span> people</p>
                        <p class="guest-indicator">Note: Guest registrations (negative IDs) cannot be tagged or added to organizations. Convert them to people records first.</p>
                    
                        <!-- Tag Section -->
                        <div class="action-section">
                            <h4>1. Tag Selected People</h4>
                            <form id="tagForm" onsubmit="return performBulkAction(event, 'tag')">
                                <input type="hidden" name="action" value="tag">
                                <input type="hidden" name="selected_people" id="tag_selected_people">
                                <input type="hidden" name="p1" value='""" + org_ids + """'>
                                <label>Tag Name: </label>
                                <input type="text" name="tag_name" required style="width: 300px;">
                                <button type="submit" class="btn btn-primary">Apply Tag</button>
                            </form>
                        </div>
                    
                        <!-- Add to Organization Section -->
                        <div class="action-section">
                            <h4>2. Add to Organization/Subgroup</h4>
                            <form id="orgForm" onsubmit="return performBulkAction(event, 'add_to_org')">
                                <input type="hidden" name="action" value="add_to_org">
                                <input type="hidden" name="selected_people" id="org_selected_people">
                                <input type="hidden" name="p1" value='""" + org_ids + """'>
                                <label>Organization: </label>
                                <select name="target_org_id" required style="width: 300px;">
                                    <option value="">Select Organization...</option>
                """)
            
                # Add all active organizations grouped by program
                # Check if only one org was selected for the registration data
                selected_org_id = None
                if len(org_id_list) == 1 and org_id_list[0].isdigit():
                    selected_org_id = int(org_id_list[0])
            
                # Group organizations by program
                current_program = None
                for org in all_orgs:
                    # Start new optgroup if program changes
                    if org.ProgramName != current_program:
                        if current_program is not None:
                            print('</optgroup>')
                        print('<optgroup label="{0}">'.format(org.ProgramName))
                        current_program = org.ProgramName
                    
                    if selected_org_id and org.OrganizationId == selected_org_id:
                        print('<option value="{0}" selected>{1} (current involvement)</option>'.format(org.OrganizationId, org.OrganizationName))
                    else:
                        print('<option value="{0}">{1}</option>'.format(org.OrganizationId, org.OrganizationName))
                
                # Close last optgroup
                if current_program is not None:
                    print('</optgroup>')
            
                print("""
                                </select><br><br>
                                <label>Subgroup (optional): </label>
                                <input type="text" name="subgroup" style="width: 300px;" placeholder="e.g., Team A, Morning Session">
                                <button type="submit" class="btn btn-success">Add to Organization</button>
                                <p style="margin-top: 10px; font-size: 12px; color: #666;">
                                    <em><strong>How it works:</strong><br>
                                    • People will be added to the organization (if not already members)<br>
                                    • If subgroup is specified, they'll be added to that subgroup<br>
                                    • Subgroups are created automatically if they don't exist<br>
                                    • Subgroup membership is tracked in TouchPoint and as "OrgSubgroup_[OrgId]" extra value</em>
                                </p>
                            </form>
                        </div>
                    
                        <!-- Save to Extra Value Section -->
                        <div class="action-section">
                            <h4>3. Save Question Answer to Extra Value</h4>
                            <form id="extraValueForm" onsubmit="return performBulkAction(event, 'save_to_extravalue')">
                                <input type="hidden" name="action" value="save_to_extravalue">
                                <input type="hidden" name="selected_people" id="ev_selected_people">
                                <input type="hidden" name="p1" value='""" + org_ids + """'>
                                <label>Question Column: </label>
                                <span style="font-size: 12px; color: #666;">(See Question Reference table above for full question text)</span><br>
                                <select name="question_col" id="question_col_select" required style="width: 300px;" onchange="updateValuePreview()">
                                    <option value="">Select Question...</option>
                """)
            
                # Add question columns
                for idx, label in enumerate(question_headers):
                    col_name = "Q{}".format(idx + 1)
                    print('<option value="{0}">{0}</option>'.format(col_name))
            
                print("""
                                </select><br><br>
                                <label>Extra Value Field Name: </label>
                                <input type="text" name="extravalue_field" required style="width: 300px;" placeholder="e.g., PreferredTime, DietaryRestriction">
                                <br><br>
                            
                                <!-- Value Options -->
                                <div style="background: #f9f9f9; padding: 15px; border-radius: 5px; margin: 10px 0;">
                                    <label style="font-weight: bold;">Value Options:</label><br>
                                    <label style="font-weight: normal;">
                                        <input type="radio" name="value_option" value="original" checked onchange="toggleGlobalValue()">
                                        Use original answers from registration
                                    </label><br>
                                    <label style="font-weight: normal;">
                                        <input type="radio" name="value_option" value="global" onchange="toggleGlobalValue()">
                                        Set a global value for all selected people
                                    </label>
                                
                                    <div id="globalValueInput" style="display: none; margin-top: 10px;">
                                        <label>Global Value: </label>
                                        <input type="text" id="global_value" name="global_value" style="width: 300px;" placeholder="Enter value to apply to all selected people">
                                    </div>
                                
                                    <div id="valuePreview" style="margin-top: 10px; font-size: 12px; color: #666;"></div>
                                </div>
                            
                                <input type="hidden" name="use_global_value" id="use_global_value" value="0">
                                <button type="submit" class="btn btn-warning">Save to Extra Value</button>
                            </form>
                        </div>
                """)
            
                # Add Admin-only section for updating person fields
                if model.UserIsInRole("Admin"):
                    print("""
                        <!-- Admin-Only Section -->
                        <div class="action-section" style="border: 2px solid #ff9800;">
                            <h4>4. Update Person Fields <span style="color: #ff9800;">(Admin Only)</span></h4>
                            <p style="font-size: 12px; color: #666;">
                                <em>Update standard person fields like Campus, Member Status, Position in Family, etc.</em>
                            </p>
                            <form id="personFieldForm" onsubmit="return performBulkAction(event, 'update_person_field')">
                                <input type="hidden" name="action" value="update_person_field">
                                <input type="hidden" name="selected_people" id="pf_selected_people">
                                <input type="hidden" name="p1" value='""" + org_ids + """'>
                            
                                <label>Field to Update: </label>
                                <select name="field_name" id="field_name_select" required style="width: 300px;" onchange="updateFieldOptions()">
                                    <option value="">Select Field...</option>
                                    <optgroup label="General Fields">
                                        <option value="Title">Title (Mr., Mrs., Dr., etc.)</option>
                                        <option value="Suffix">Suffix (Jr., Sr., III, etc.)</option>
                                        <option value="MiddleName">Middle Name</option>
                                        <option value="MaidenName">Maiden Name</option>
                                        <option value="NickName">Nick Name</option>
                                    </optgroup>
                                    <optgroup label="Status Fields">
                                        <option value="CampusId">Campus</option>
                                        <option value="MemberStatusId">Member Status</option>
                                        <option value="PositionInFamilyId">Position in Family</option>
                                        <option value="GenderId">Gender</option>
                                        <option value="MaritalStatusId">Marital Status</option>
                                    </optgroup>
                                    <optgroup label="Contact Preferences">
                                        <option value="DoNotCallFlag">Do Not Call</option>
                                        <option value="DoNotMailFlag">Do Not Mail</option>
                                        <option value="DoNotVisitFlag">Do Not Visit</option>
                                        <option value="DoNotPublishPhones">Do Not Publish Phones</option>
                                    </optgroup>
                                    <optgroup label="Other Fields">
                                        <option value="SchoolOther">School</option>
                                        <option value="EmployerOther">Employer</option>
                                        <option value="OccupationOther">Occupation</option>
                                        <option value="AltName">Alternate Name</option>
                                    </optgroup>
                                </select><br><br>
                            
                                <!-- Value Options for all field types -->
                                <div id="fieldValueOptions" style="background: #f9f9f9; padding: 15px; border-radius: 5px; margin: 10px 0;">
                                    <label style="font-weight: bold;">Value Options:</label><br>
                                
                                    <!-- Options for text fields -->
                                    <div id="textFieldOptions" style="display: none;">
                                        <label style="font-weight: normal;">
                                            <input type="radio" name="field_value_option" value="global" checked onchange="toggleFieldValueOption()">
                                            Set a global value for all selected people
                                        </label><br>
                                        <label style="font-weight: normal;">
                                            <input type="radio" name="field_value_option" value="current" onchange="toggleFieldValueOption()">
                                            Use each person's current value
                                        </label>
                                    </div>
                                
                                    <!-- For lookup fields, always use global value -->
                                    <div id="lookupFieldNote" style="display: none; font-size: 12px; color: #666;">
                                        <em>Select a value to apply to all selected people</em>
                                    </div>
                                
                                    <!-- For checkbox fields, always use global value -->
                                    <div id="checkboxFieldNote" style="display: none; font-size: 12px; color: #666;">
                                        <em>This setting will be applied to all selected people</em>
                                    </div>
                                
                                    <div id="fieldCurrentValuePreview" style="margin-top: 10px; font-size: 12px; color: #666;"></div>
                                </div>
                            
                                <div id="fieldValueInput">
                                    <label>New Value: </label>
                                    <input type="text" name="field_value" id="field_value" style="width: 300px;" placeholder="Enter new value">
                                </div>
                            
                                <div id="fieldLookupInput" style="display: none;">
                                    <label>New Value: </label>
                                    <select name="field_lookup_value" id="field_lookup_value" style="width: 300px;">
                                        <option value="">Loading options...</option>
                                    </select>
                                </div>
                            
                                <div id="fieldCheckboxInput" style="display: none;">
                                    <label>
                                        <input type="checkbox" name="field_checkbox_value" id="field_checkbox_value" value="1">
                                        <span id="checkboxLabel">Enable this option</span>
                                    </label>
                                </div>
                            
                                <input type="hidden" name="use_current_value" id="use_current_value" value="0">
                            
                                <br><br>
                                <button type="submit" class="btn btn-warning">Update Person Field</button>
                            
                                <p style="margin-top: 10px; font-size: 12px; color: #666;">
                                    <em><strong>Warning:</strong> This will directly update person records. Changes cannot be easily undone.</em>
                                </p>
                            </form>
                        </div>
                    """)
            
                print("""
                    </div>
                
                    <!-- Grid Wrapper -->
                    <div class="grid-wrapper">
                        <!-- Row Grouping Panel -->
                        <div id="rowGroupPanel" class="ag-row-group-panel ag-theme-alpine" style="display: none;"></div>
                    
                        <!-- Data Grid -->
                        <div id="myGrid" class="ag-theme-alpine"></div>
                    </div>
                </div>
            
                <script>
                // Grid data and question mapping
                let gridData = """ + json.dumps(grid_data) + """;
                let questionMapping = """ + json.dumps(question_mapping) + """;
                
                // Pre-loaded lookup values (Admin only)
                let preloadedLookupValues = """ + json.dumps(lookup_values) + """;
            
                // Define admin functions globally but they'll only be called if admin section exists
                // These will be overridden with actual implementations below if admin section exists
            
                console.log('Grid data loaded:', gridData.length, 'records');
                console.log('Question mapping:', questionMapping);
                console.log('Preloaded lookup values:', Object.keys(preloadedLookupValues));
            
                // Get form action URL dynamically
                function getPyScriptFormUrl() {
                    const currentPath = window.location.pathname;
                    const currentUrl = window.location.href;
                
                    // Handle both /PyScript/ and /PyScriptForm/ paths
                    if (currentPath.includes('/PyScriptForm/')) {
                        return currentUrl; // Already correct
                    } else if (currentPath.includes('/PyScript/')) {
                        return currentUrl.replace('/PyScript/', '/PyScriptForm/');
                    } else {
                        // Fallback - construct from scratch
                        const scriptName = currentPath.split('/').pop().split('?')[0];
                        return '/PyScriptForm/' + scriptName + window.location.search;
                    }
                }
            
                // Column definitions with Q1, Q2 headers but full tooltips
                const columnDefs = [
                    { 
                        field: 'checkbox',
                        headerName: '',
                        width: 50,
                        headerCheckboxSelection: true,
                        checkboxSelection: true,
                        headerCheckboxSelectionFilteredOnly: true
                    },
                    { 
                        field: 'PeopleId', 
                        headerName: 'ID', 
                        width: 80,
                        filter: 'agNumberColumnFilter',
                        floatingFilter: true,
                        cellRenderer: function(params) {
                            if (params.value < 0) {
                                return '<span style="color: #999;">' + params.value + ' (guest)</span>';
                            }
                            return params.value;
                        }
                    },
                    { 
                        field: 'FirstName', 
                        headerName: 'First Name', 
                        width: 120,
                        filter: 'agTextColumnFilter',
                        floatingFilter: true
                    },
                    { 
                        field: 'LastName', 
                        headerName: 'Last Name', 
                        width: 120,
                        filter: 'agTextColumnFilter',
                        floatingFilter: true
                    },
                    { 
                        field: 'EmailAddress', 
                        headerName: 'Email', 
                        width: 200,
                        filter: 'agTextColumnFilter',
                        floatingFilter: true
                    },
                    { 
                        field: 'Phone', 
                        headerName: 'Phone', 
                        width: 120,
                        filter: 'agTextColumnFilter',
                        floatingFilter: true
                    },
                    { 
                        field: 'OrganizationName', 
                        headerName: 'Organization', 
                        width: 200,
                        filter: 'agTextColumnFilter',
                        floatingFilter: true
                    }
                ];
            
                // Add question columns with Q1, Q2 headers - in order!
                const sortedQuestionCols = Object.keys(questionMapping).sort(function(a, b) {
                    // Extract numbers from Q1, Q2, etc.
                    const numA = parseInt(a.replace('Q', ''));
                    const numB = parseInt(b.replace('Q', ''));
                    return numA - numB;
                });
            
                sortedQuestionCols.forEach(function(qCol) {
                    columnDefs.push({
                        field: qCol,
                        headerName: qCol, // Use Q1, Q2, etc.
                        headerTooltip: questionMapping[qCol], // Full question on hover
                        width: 150,
                        filter: 'agTextColumnFilter',
                        floatingFilter: true,
                        cellRenderer: function(params) {
                            if (!params.value || params.value === '') {
                                return '<span style="color: #ccc;">-</span>';
                            }
                            // Create a div to safely display content
                            const div = document.createElement('div');
                            // Unescape the value for display
                            let displayValue = params.value;
                            try {
                                displayValue = displayValue.replace(/\\\\n/g, '\\n')
                                                         .replace(/\\\\r/g, '\\r')
                                                         .replace(/\\\\t/g, '\\t')
                                                         .replace(/\\\\"/g, '"')
                                                         .replace(/\\\\\\\\/g, '\\\\');
                            } catch(e) {
                                console.error('Error unescaping value:', e);
                            }
                            div.textContent = displayValue;
                            div.title = displayValue; // Tooltip with full value
                            return div.innerHTML;
                        }
                    });
                });
            
                // Grid options with pagination support and sidebar
                const gridOptions = {
                    columnDefs: columnDefs,
                    rowData: gridData,
                    defaultColDef: {
                        sortable: true,
                        resizable: true,
                        enableRowGroup: true,  // Allow grouping on all columns
                        enableValue: true      // Allow aggregation on all columns
                    },
                    rowSelection: 'multiple',
                    floatingFilter: true, // Enable column filters
                    onSelectionChanged: onSelectionChanged,
                    getRowNodeId: function(data) {
                        return String(data.PeopleId);
                    },
                    // Sidebar configuration - simplified for better compatibility
                    sideBar: {
                        toolPanels: [
                            {
                                id: 'columns',
                                labelDefault: 'Columns',
                                labelKey: 'columns',
                                iconKey: 'columns',
                                toolPanel: 'agColumnsToolPanel'
                            },
                            {
                                id: 'filters',
                                labelDefault: 'Filters',
                                labelKey: 'filters',
                                iconKey: 'filter',
                                toolPanel: 'agFiltersToolPanel'
                            }
                        ],
                        defaultToolPanel: null
                    },
                    // Row grouping configuration
                    groupDefaultExpanded: 1,
                    autoGroupColumnDef: {
                        headerName: 'Group',
                        minWidth: 200,
                        cellRendererParams: {
                            suppressCount: false
                        }
                    },
                    // Enable row group panel - set to 'never' initially
                    rowGroupPanelShow: 'never',
                    suppressAggFuncInHeader: true,
                    // Pagination settings from config
                    pagination: """ + ("true" if Config.ENABLE_PAGINATION else "false") + """,
                    paginationPageSize: """ + str(Config.PAGE_SIZE) + """,
                    domLayout: 'normal'  // Use normal layout to respect height setting
                };
            
                // Initialize grid
                document.addEventListener('DOMContentLoaded', function() {
                    const gridDiv = document.querySelector('#myGrid');
                    new agGrid.Grid(gridDiv, gridOptions);
                
                    // Show row group panel based on config
                    const rowGroupPanel = document.getElementById('rowGroupPanel');
                    if (rowGroupPanel && gridOptions.rowGroupPanelShow === 'always') {
                        rowGroupPanel.style.display = 'block';
                    }
                });
            
                // Selection changed handler
                function onSelectionChanged() {
                    const selectedRows = gridOptions.api.getSelectedRows();
                    document.getElementById('selectedCount').textContent = selectedRows.length;
                
                    // Update value preview
                    updateValuePreview();
                }
            
                // Quick filter
                function onQuickFilter() {
                    const filterValue = document.getElementById('quickFilter').value;
                    gridOptions.api.setQuickFilter(filterValue);
                }
            
                // Clear all filters
                function clearFilters() {
                    document.getElementById('quickFilter').value = '';
                    gridOptions.api.setQuickFilter('');
                    gridOptions.api.setFilterModel(null);
                }
            
                // Toggle sidebar
                function toggleSidebar() {
                    if (!gridOptions.api) return;
                
                    const isVisible = gridOptions.api.isSideBarVisible();
                    gridOptions.api.setSideBarVisible(!isVisible);
                
                    // If opening, default to columns panel
                    if (!isVisible) {
                        gridOptions.api.openToolPanel('columns');
                    }
                }
            
                // Toggle global value input
                function toggleGlobalValue() {
                    const globalOption = document.querySelector('input[name="value_option"]:checked').value;
                    const globalInput = document.getElementById('globalValueInput');
                    const useGlobal = document.getElementById('use_global_value');
                
                    if (globalOption === 'global') {
                        globalInput.style.display = 'block';
                        useGlobal.value = '1';
                    } else {
                        globalInput.style.display = 'none';
                        useGlobal.value = '0';
                    }
                
                    updateValuePreview();
                }
            
                // Update value preview
                function updateValuePreview() {
                    const selectedRows = gridOptions.api.getSelectedRows();
                    const questionCol = document.getElementById('question_col_select').value;
                    const previewDiv = document.getElementById('valuePreview');
                
                    if (!questionCol || selectedRows.length === 0) {
                        previewDiv.innerHTML = '';
                        return;
                    }
                
                    const globalOption = document.querySelector('input[name="value_option"]:checked').value;
                
                    if (globalOption === 'global') {
                        previewDiv.innerHTML = '<strong>Note:</strong> The global value will be applied to all ' + selectedRows.length + ' selected people.';
                    } else {
                        // Show sample of original values
                        let sampleValues = [];
                        let nullCount = 0;
                    
                        selectedRows.slice(0, 5).forEach(row => {
                            if (row[questionCol]) {
                                // Unescape for display
                                let displayValue = row[questionCol];
                                try {
                                    displayValue = displayValue.replace(/\\\\n/g, ' ')
                                                             .replace(/\\\\r/g, ' ')
                                                             .replace(/\\\\t/g, ' ')
                                                             .replace(/\\\\"/g, '"')
                                                             .replace(/\\\\\\\\/g, '\\\\');
                                } catch(e) {}
                                sampleValues.push('"' + displayValue + '"');
                            } else {
                                nullCount++;
                            }
                        });
                    
                        let preview = '<strong>Sample values:</strong> ';
                        if (sampleValues.length > 0) {
                            preview += sampleValues.join(', ');
                            if (selectedRows.length > 5) {
                                preview += ', ...';
                            }
                        }
                        if (nullCount > 0) {
                            preview += ' (' + nullCount + ' empty values)';
                        }
                    
                        previewDiv.innerHTML = preview;
                    }
                }
            
                // Perform bulk action
                function performBulkAction(event, actionType) {
                    event.preventDefault();
                
                    const selectedRows = gridOptions.api.getSelectedRows();
                    if (selectedRows.length === 0) {
                        alert('Please select at least one person.');
                        return false;
                    }
                
                    // Filter out negative IDs (guests) for certain actions
                    let peopleIds = selectedRows.map(row => row.PeopleId);
                    if (actionType === 'tag' || actionType === 'add_to_org' || actionType === 'save_to_extravalue') {
                        const validPeopleIds = peopleIds.filter(id => id > 0);
                        if (validPeopleIds.length === 0) {
                            alert('This action cannot be performed on guest registrations. Please select people with valid IDs.');
                            return false;
                        }
                        // Use all selected IDs but server will skip guests
                    }
                
                    const peopleIdsStr = peopleIds.join(',');
                
                    if (actionType === 'tag') {
                        document.getElementById('tag_selected_people').value = peopleIdsStr;
                    } else if (actionType === 'add_to_org') {
                        document.getElementById('org_selected_people').value = peopleIdsStr;
                    } else if (actionType === 'save_to_extravalue') {
                        document.getElementById('ev_selected_people').value = peopleIdsStr;
                    
                        // Collect answers for selected question
                        const questionCol = document.getElementById('question_col_select').value;
                        if (!questionCol) {
                            alert('Please select a question column.');
                            return false;
                        }
                    
                        // Check if using global value
                        const useGlobal = document.getElementById('use_global_value').value;
                        if (useGlobal === '1') {
                            const globalValue = document.getElementById('global_value').value;
                            if (!globalValue) {
                                alert('Please enter a global value to apply.');
                                return false;
                            }
                        }
                    
                        const answers = {};
                        selectedRows.forEach(row => {
                            if (row[questionCol]) {
                                answers[row.PeopleId] = row[questionCol];
                            }
                        });
                    
                        // Add answers as hidden field
                        const answersInput = document.createElement('input');
                        answersInput.type = 'hidden';
                        answersInput.name = 'answers_json';
                        answersInput.value = JSON.stringify(answers);
                        event.target.appendChild(answersInput);
                    } else if (actionType === 'update_person_field') {
                        document.getElementById('pf_selected_people').value = peopleIdsStr;
                    
                        // Get field name and value
                        const fieldName = document.getElementById('field_name_select').value;
                        if (!fieldName) {
                            alert('Please select a field to update.');
                            return false;
                        }
                    
                        // Check if using current value (for text fields)
                        const useCurrentValue = document.getElementById('use_current_value').value;
                    
                        // Get the appropriate value based on field type
                        const lookupFields = ['CampusId', 'MemberStatusId', 'PositionInFamilyId', 'GenderId', 'MaritalStatusId'];
                        const checkboxFields = ['DoNotCallFlag', 'DoNotMailFlag', 'DoNotVisitFlag', 'DoNotPublishPhones'];
                    
                        let fieldValue = '';
                        if (useCurrentValue === '1') {
                            // Using current value - no need to get a value
                            fieldValue = '';
                        } else if (lookupFields.includes(fieldName)) {
                            fieldValue = document.getElementById('field_lookup_value').value;
                            if (!fieldValue) {
                                alert('Please select a value for the field.');
                                return false;
                            }
                        } else if (checkboxFields.includes(fieldName)) {
                            fieldValue = document.getElementById('field_checkbox_value').checked ? '1' : '0';
                        } else {
                            fieldValue = document.getElementById('field_value').value;
                            if (!fieldValue) {
                                alert('Please enter a value for the field.');
                                return false;
                            }
                        }
                    
                        // Add field value as hidden input
                        const valueInput = document.createElement('input');
                        valueInput.type = 'hidden';
                        valueInput.name = 'field_value_final';
                        valueInput.value = fieldValue;
                        event.target.appendChild(valueInput);
                    }
                
                    // Show loading indicator
                    const submitBtn = event.target.querySelector('button[type="submit"]');
                    const originalText = submitBtn.textContent;
                    submitBtn.textContent = 'Processing...';
                    submitBtn.disabled = true;
                
                    // Submit form
                    const formData = new FormData(event.target);
                    const submitUrl = getPyScriptFormUrl();
                
                    fetch(submitUrl, {
                        method: 'POST',
                        body: formData
                    })
                    .then(response => response.text())
                    .then(html => {
                        // Extract and display alerts
                        const tempDiv = document.createElement('div');
                        tempDiv.innerHTML = html;
                    
                        const alerts = tempDiv.querySelectorAll('.alert');
                        if (alerts.length > 0) {
                            const container = document.querySelector('.manager-container');
                            alerts.forEach(alert => {
                                container.insertBefore(alert.cloneNode(true), container.firstChild);
                            });
                        
                            // Clear form
                            event.target.reset();
                        
                            // Auto-hide success messages
                            setTimeout(() => {
                                document.querySelectorAll('.alert-success').forEach(alert => {
                                    alert.style.display = 'none';
                                });
                            }, 5000);
                        }
                    
                        // Restore button
                        submitBtn.textContent = originalText;
                        submitBtn.disabled = false;
                    })
                    .catch(error => {
                        console.error('Error:', error);
                        alert('An error occurred: ' + error.message);
                        submitBtn.textContent = originalText;
                        submitBtn.disabled = false;
                    });
                
                    return false;
                }
            
                // Export selected with full question headers
                function exportSelected() {
                    const selectedRows = gridOptions.api.getSelectedRows();
                    if (selectedRows.length === 0) {
                        alert('Please select at least one person to export.');
                        return;
                    }
                
                    // Check export limit
                    const maxExport = """ + str(Config.MAX_EXPORT_ROWS) + """;
                    if (selectedRows.length > maxExport) {
                        alert('Export limited to ' + maxExport + ' rows. Only the first ' + maxExport + ' rows will be exported.');
                    }
                
                    // Create CSV with full headers
                    const headers = ['ID', 'First Name', 'Last Name', 'Email', 'Phone', 'Organization'];
                
                    // Add full question headers - in sorted order
                    const sortedQuestionCols = Object.keys(questionMapping).sort(function(a, b) {
                        const numA = parseInt(a.replace('Q', ''));
                        const numB = parseInt(b.replace('Q', ''));
                        return numA - numB;
                    });
                
                    sortedQuestionCols.forEach(qCol => {
                        headers.push(questionMapping[qCol]);
                    });
                
                    const csvRows = [headers.join(',')];
                
                    // Limit rows to export
                    const rowsToExport = selectedRows.slice(0, maxExport);
                
                    rowsToExport.forEach(row => {
                        const values = [
                            row.PeopleId,
                            '"' + (row.FirstName || '').replace(/"/g, '""') + '"',
                            '"' + (row.LastName || '').replace(/"/g, '""') + '"',
                            '"' + (row.EmailAddress || '').replace(/"/g, '""') + '"',
                            '"' + (row.Phone || '').replace(/"/g, '""') + '"',
                            '"' + (row.OrganizationName || '').replace(/"/g, '""') + '"'
                        ];
                    
                        // Add question answers - in sorted order
                        sortedQuestionCols.forEach(qCol => {
                            let value = row[qCol] || '';
                            // Unescape for export
                            try {
                                value = value.replace(/\\\\n/g, '\\n')
                                           .replace(/\\\\r/g, '\\r')
                                           .replace(/\\\\t/g, '\\t')
                                           .replace(/\\\\"/g, '"')
                                           .replace(/\\\\\\\\/g, '\\\\');
                            } catch(e) {}
                            values.push('"' + value.toString().replace(/"/g, '""') + '"');
                        });
                    
                        csvRows.push(values.join(','));
                    });
                
                    // Download CSV
                    const csvContent = csvRows.join('\\n');
                    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                    const link = document.createElement('a');
                    const url = URL.createObjectURL(blob);
                    link.setAttribute('href', url);
                    link.setAttribute('download', '""" + Config.EXPORT_FILENAME + """');
                    link.style.visibility = 'hidden';
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                }
            
                // Export to Excel function
                function exportToExcel() {
                    if (!gridOptions.api) {
                        alert('Grid not initialized');
                        return;
                    }
                
                    // Get export parameters
                    const params = {
                        fileName: 'registration_data_' + new Date().toISOString().slice(0, 10),
                        sheetName: 'Registration Data',
                        // Export all rows or just selected
                        onlySelected: false,
                        // Include headers with full question text
                        processCellCallback: function(params) {
                            // Handle special formatting if needed
                            return params.value;
                        },
                        // Custom header names
                        processHeaderCallback: function(params) {
                            // For question columns, use full question text
                            if (params.column.colId.startsWith('Q') && questionMapping[params.column.colId]) {
                                return questionMapping[params.column.colId];
                            }
                            return params.column.colDef.headerName;
                        }
                    };
                
                    // Export to Excel
                    gridOptions.api.exportDataAsExcel(params);
                }
            
                // Update field options based on selected field (Admin only)
                window.updateFieldOptions = function() {
                    const fieldName = document.getElementById('field_name_select').value;
                    const textInput = document.getElementById('fieldValueInput');
                    const lookupInput = document.getElementById('fieldLookupInput');
                    const checkboxInput = document.getElementById('fieldCheckboxInput');
                    const checkboxLabel = document.getElementById('checkboxLabel');
                    const valueOptions = document.getElementById('fieldValueOptions');
                    const textFieldOptions = document.getElementById('textFieldOptions');
                    const lookupFieldNote = document.getElementById('lookupFieldNote');
                    const checkboxFieldNote = document.getElementById('checkboxFieldNote');
                
                    // Hide all inputs and options first
                    textInput.style.display = 'none';
                    lookupInput.style.display = 'none';
                    checkboxInput.style.display = 'none';
                    valueOptions.style.display = 'none';
                    textFieldOptions.style.display = 'none';
                    lookupFieldNote.style.display = 'none';
                    checkboxFieldNote.style.display = 'none';
                
                    if (!fieldName) return;
                
                    // Show value options section
                    valueOptions.style.display = 'block';
                
                    // Determine input type based on field
                    const lookupFields = ['CampusId', 'MemberStatusId', 'PositionInFamilyId', 'GenderId', 'MaritalStatusId'];
                    const checkboxFields = ['DoNotCallFlag', 'DoNotMailFlag', 'DoNotVisitFlag', 'DoNotPublishPhones'];
                
                    if (lookupFields.includes(fieldName)) {
                        // Show lookup dropdown
                        lookupInput.style.display = 'block';
                        lookupFieldNote.style.display = 'block';
                    
                        // Load options from preloaded data (no AJAX needed)
                        const lookupSelect = document.getElementById('field_lookup_value');
                        lookupSelect.innerHTML = '<option value="">Select...</option>';
                        
                        console.log('Loading lookup values for:', fieldName);
                        console.log('Available preloaded data:', Object.keys(preloadedLookupValues));
                        
                        if (preloadedLookupValues[fieldName]) {
                            const values = preloadedLookupValues[fieldName];
                            console.log('Found', values.length, 'values for', fieldName);
                            
                            if (values.length > 0) {
                                values.forEach(item => {
                                    const option = document.createElement('option');
                                    option.value = item.id;
                                    option.textContent = item.name;
                                    lookupSelect.appendChild(option);
                                });
                            } else {
                                lookupSelect.innerHTML = '<option value="">No values available</option>';
                            }
                        } else {
                            console.error('No preloaded data for field:', fieldName);
                            lookupSelect.innerHTML = '<option value="">No values available</option>';
                        }
                    
                    } else if (checkboxFields.includes(fieldName)) {
                        // Show checkbox
                        checkboxInput.style.display = 'block';
                        checkboxFieldNote.style.display = 'block';
                    
                        // Update checkbox label
                        const labels = {
                            'DoNotCallFlag': 'Do Not Call',
                            'DoNotMailFlag': 'Do Not Mail',
                            'DoNotVisitFlag': 'Do Not Visit',
                            'DoNotPublishPhones': 'Do Not Publish Phone Numbers'
                        };
                        checkboxLabel.textContent = labels[fieldName] || 'Enable this option';
                    
                    } else {
                        // Show text input with options
                        textFieldOptions.style.display = 'block';
                        textInput.style.display = 'block';
                    
                        // Reset to global value option by default
                        document.querySelector('input[name="field_value_option"][value="global"]').checked = true;
                        toggleFieldValueOption();
                    }
                
                    // Update the preview for current values
                    updateFieldValuePreview();
                };
            
                // Toggle between global value and current value for text fields
                window.toggleFieldValueOption = function() {
                    const fieldOption = document.querySelector('input[name="field_value_option"]:checked');
                    const textInput = document.getElementById('fieldValueInput');
                    const useCurrentValue = document.getElementById('use_current_value');
                
                    if (fieldOption && fieldOption.value === 'current') {
                        textInput.style.display = 'none';
                        useCurrentValue.value = '1';
                    } else {
                        textInput.style.display = 'block';
                        useCurrentValue.value = '0';
                    }
                
                    updateFieldValuePreview();
                };
            
                // Update preview of current field values
                window.updateFieldValuePreview = function() {
                    const selectedRows = gridOptions.api.getSelectedRows();
                    const fieldName = document.getElementById('field_name_select').value;
                    const previewDiv = document.getElementById('fieldCurrentValuePreview');
                    const fieldOption = document.querySelector('input[name="field_value_option"]:checked');
                
                    if (!fieldName || selectedRows.length === 0) {
                        previewDiv.innerHTML = '';
                        return;
                    }
                
                    if (fieldOption && fieldOption.value === 'current') {
                        previewDiv.innerHTML = '<strong>Note:</strong> Each person\\\'s current ' + fieldName + ' value will be preserved. This is useful for appending or modifying existing values.';
                    } else {
                        previewDiv.innerHTML = '<strong>Note:</strong> The value will be applied to all ' + selectedRows.length + ' selected people.';
                    }
                };
                </script>
                """)
            
                # Show debug panel for admins
                show_debug_panel()
            
                # Add Mermaid visualization for SuperAdmin users only
                if Config.ENABLE_MERMAID and model.UserIsInRole(Config.MERMAID_ROLE):
                    print("""
                    <!-- Code Structure Visualization for SuperAdmin -->
                    <details class="mermaid-output" style="background: #f5f5f5; padding: 20px; border-radius: 8px; margin: 20px 0;">
                        <summary style="cursor: pointer; font-size: 16px; font-weight: bold; color: #1976d2; padding: 10px;">
                            📊 Code Structure Visualization (Click to expand)
                        </summary>
                        <div style="margin-top: 20px;">
                            <h4>Registration Data Manager - Code Structure Visualization</h4>
                            <div class="mermaid">
                            graph TD
                                A[User Selects Organizations] -->|p1 parameter| B[Load Registration Data]
                                B --> C{Process Data}
                            
                                C -->|Questions| D[Extract Unique Questions]
                                C -->|People| E[Build Person Records]
                                C -->|Answers| F[Map Answers to Questions]
                            
                                D --> G[Create Column Mappings]
                                E --> G
                                F --> G
                            
                                G --> H[Interactive Grid Display]
                            
                                H -->|User Actions| I{Bulk Operations}
                            
                                I -->|Tag| J[Create/Apply Tags]
                                I -->|Organization| K[Add to Org/Subgroup]
                                I -->|Extra Value| L[Save to Person Fields]
                                I -->|Export| M[Generate CSV]
                            
                                J --> N[TouchPoint API:<br/>model.CreateQueryTag]
                                K --> O[TouchPoint APIs:<br/>model.AddMemberToOrg<br/>model.AddSubGroup]
                                L --> P[TouchPoint API:<br/>model.AddExtraValueText]
                            
                                style A fill:#e1f5fe,stroke:#01579b,stroke-width:3px
                                style H fill:#f3e5f5,stroke:#4a148c,stroke-width:3px
                                style I fill:#fff3e0,stroke:#e65100,stroke-width:3px
                                style N fill:#e8f5e9,stroke:#1b5e20,stroke-width:2px
                                style O fill:#e8f5e9,stroke:#1b5e20,stroke-width:2px
                                style P fill:#e8f5e9,stroke:#1b5e20,stroke-width:2px
                            
                                subgraph "Data Flow"
                                    B
                                    C
                                    D
                                    E
                                    F
                                    G
                                end
                            
                                subgraph "User Interface"
                                    H
                                    I
                                    J
                                    K
                                    L
                                    M
                                end
                            
                                subgraph "TouchPoint Integration"
                                    N
                                    O
                                    P
                                end
                            </div>
                        
                            <h4 style="margin-top: 20px;">Process Flow Details</h4>
                            <div class="mermaid">
                            sequenceDiagram
                                participant U as User
                                participant S as Script
                                participant DB as TouchPoint DB
                                participant API as TouchPoint API
                            
                                U->>S: Select Organizations
                                S->>DB: Query Registration Data
                                DB-->>S: Registration Records
                                S->>DB: Query Questions
                                DB-->>S: Question Definitions
                            
                                Note over S: Process & Transform Data
                            
                                S->>U: Display Interactive Grid
                            
                                U->>S: Select Records
                                U->>S: Choose Bulk Action
                            
                                alt Tag Action
                                    S->>API: CreateQueryTag
                                    API-->>S: Tag Created
                                else Org Action
                                    S->>API: AddMemberToOrg
                                    API-->>S: Member Added
                                    S->>API: AddSubGroup
                                    API-->>S: Subgroup Assigned
                                else Extra Value
                                    S->>API: AddExtraValueText
                                    API-->>S: Value Saved
                                end
                            
                                S->>U: Show Success Message
                            </div>
                        </div>
                    </details>
                
                    <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
                    <script>
                        mermaid.initialize({ startOnLoad: true });
                    
                        // Re-render mermaid when details is opened
                        document.addEventListener('DOMContentLoaded', function() {
                            const details = document.querySelector('details.mermaid-output');
                            if (details) {
                                details.addEventListener('toggle', function() {
                                    if (this.open) {
                                        mermaid.init();
                                    }
                                });
                            }
                        });
                    </script>
                
                    <style>
                    .mermaid-output {
                        border: 2px solid #ddd;
                        margin-top: 40px;
                    }
                    .mermaid-output summary {
                        outline: none;
                    }
                    .mermaid-output summary:hover {
                        background-color: #e3f2fd;
                        border-radius: 5px;
                    }
                    .mermaid-output h4 {
                        color: #1976d2;
                        margin-bottom: 15px;
                    }
                    </style>
                    """)
        
        except Exception as e:
            print_error("Main Process", e, show_traceback=True)
            debug_print("Fatal error occurred", {
                "error": str(e),
                "org_ids": org_ids
            })
# Call main function
main()
