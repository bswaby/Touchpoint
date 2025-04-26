#####################################################################
#### MINISTRY STRUCTURE
#####################################################################
# This dashboard displays the hierarchical structure of your church's
# Programs, Divisions, and Involvements with key metrics and filters.
# Features:
# - Expandable/collapsible tree view
# - Metrics for members, previous members, visitors, and meetings
# - Filter options for active/inactive involvements and by program
# - Filter by creation date to track new involvements
# - Counts and summary statistics with roll-up totals
# - Support for involvements assigned to multiple divisions
#####################################################################

# written by: Ben Swaby
# email: bswaby@fbchtn.org

#roles=Admin,Finance,ManagePeople,ManageGroups

# Set page header
model.Header = "Ministry Structure"

def get_summary_stats(show_inactive=False, selected_program=None, created_after=None):
    """Generate summary statistics for the dashboard header with filters"""
    # Build the WHERE clause based on filters
    where_clause = ""
    joins = ""
    
    if not show_inactive:
        where_clause = " WHERE os.OrgStatus = 'Active'"
    
    if selected_program and selected_program != "0":
        if where_clause:
            where_clause += " AND os.ProgId = " + str(selected_program)
        else:
            where_clause = " WHERE os.ProgId = " + str(selected_program)
    
    # Add CreatedDate filter if provided
    if created_after:
        joins = " INNER JOIN Organizations o ON o.OrganizationId = os.OrgId"
        if where_clause:
            where_clause += " AND o.CreatedDate >= '" + created_after + "'"
        else:
            where_clause = " WHERE o.CreatedDate >= '" + created_after + "'"
    
    # SQL to get summary counts with filters applied
    sql_summary = """
    SELECT 
        COUNT(DISTINCT os.ProgId) AS ProgramCount,
        COUNT(DISTINCT os.DivId) AS DivisionCount,
        COUNT(DISTINCT os.OrgId) AS InvCount,
        SUM(os.Members) AS TotalMembers,
        SUM(os.Previous) AS TotalPrevious,
        SUM(os.Vistors) AS TotalVisitors,
        SUM(os.Meetings) AS TotalMeetings,
        COUNT(DISTINCT CASE WHEN os.OrgStatus = 'Active' THEN os.OrgId END) AS ActiveInvs
    FROM OrganizationStructure os
    """ + joins + where_clause
    
    try:
        return q.QuerySqlTop1(sql_summary)
    except Exception as e:
        print "<p>Error retrieving summary data: " + str(e) + "</p>"
        return None

def get_structure_data_with_related(show_inactive=False, selected_program=None, created_after=None):
    """Get hierarchical structure data with related divisions in a single query"""
    # Build the WHERE clause based on filters
    where_clause = ""
    
    if not show_inactive:
        where_clause = " WHERE os.OrgStatus = 'Active'"
    
    if selected_program and selected_program != "0":
        if where_clause:
            where_clause += " AND os.ProgId = " + str(selected_program)
        else:
            where_clause = " WHERE os.ProgId = " + str(selected_program)
    
    # Add CreatedDate filter if provided
    if created_after:
        if where_clause:
            where_clause += " AND o.CreatedDate >= '" + created_after + "'"
        else:
            where_clause = " WHERE o.CreatedDate >= '" + created_after + "'"
    
    # SQL to get structure data with related divisions and organization type in one query
    # Added OrganizationType (ot.Description) as OrgType and o.CreatedDate
    sql_structure = """
    WITH RelatedDivisions AS (
        SELECT 
            o.OrganizationId AS OrgId,
            d.Id AS RelatedDivId,
            d.Name AS RelatedDivName,
            ROW_NUMBER() OVER (PARTITION BY o.OrganizationId, d.Id ORDER BY d.Name) AS RowNum
        FROM Organizations o
        JOIN DivOrg do ON o.OrganizationId = do.OrgId
        JOIN Division d ON do.DivId = d.Id
    ),
    RelatedDivsAgg AS (
        SELECT 
            OrgId,
            STRING_AGG(CAST(RelatedDivId AS VARCHAR) + ':' + RelatedDivName, '|') AS RelatedDivs
        FROM RelatedDivisions
        GROUP BY OrgId
    )
    SELECT 
        os.ProgId,
        os.Program,
        os.DivId,
        os.Division,
        os.OrgId,
        os.Organization,
        os.OrgStatus,
        os.Members,
        os.Previous,
        os.Vistors AS Visitors,
        os.Meetings,
        rd.RelatedDivs,
        ot.Description AS OrgType,
        o.CreatedDate
    FROM OrganizationStructure os
    INNER JOIN Organizations o ON o.OrganizationId = os.OrgId
    INNER JOIN lookup.OrganizationType ot ON ot.Id = o.OrganizationTypeId
    LEFT JOIN RelatedDivsAgg rd ON os.OrgId = rd.OrgId
    """ + where_clause + """
    ORDER BY os.Program, os.Division, os.Organization
    """
    
    try:
        return q.QuerySql(sql_structure)
    except Exception as e:
        # Try a simpler query if the complex one fails
        print "<p>Error with enhanced query. Falling back to basic query.</p>"
        return get_structure_data_basic(show_inactive, selected_program, created_after)

def get_structure_data_basic(show_inactive=False, selected_program=None, created_after=None):
    """Fallback method that gets just the basic structure data"""
    # Build the WHERE clause based on filters
    where_clause = ""
    
    if not show_inactive:
        where_clause = " WHERE os.OrgStatus = 'Active'"
    
    if selected_program and selected_program != "0":
        if where_clause:
            where_clause += " AND os.ProgId = " + str(selected_program)
        else:
            where_clause = " WHERE os.ProgId = " + str(selected_program)
    
    # Add CreatedDate filter if provided
    if created_after:
        if where_clause:
            where_clause += " AND o.CreatedDate >= '" + created_after + "'"
        else:
            where_clause = " WHERE o.CreatedDate >= '" + created_after + "'"
    
    # Basic query with organization type but without related divisions
    sql_structure = """
    SELECT 
        os.ProgId,
        os.Program,
        os.DivId,
        os.Division,
        os.OrgId,
        os.Organization,
        os.OrgStatus,
        os.Members,
        os.Previous,
        os.Vistors AS Visitors,
        os.Meetings,
        ot.Description AS OrgType,
        o.CreatedDate
    FROM OrganizationStructure os
    INNER JOIN Organizations o ON o.OrganizationId = os.OrgId
    INNER JOIN lookup.OrganizationType ot ON ot.Id = o.OrganizationTypeId
    """ + where_clause + """
    ORDER BY os.Program, os.Division, os.Organization
    """
    
    try:
        return q.QuerySql(sql_structure)
    except Exception as e:
        print "<p>Error retrieving structure data: " + str(e) + "</p>"
        return []

def get_programs_for_filter():
    """Get list of programs for the filter dropdown"""
    sql_programs = """
    SELECT DISTINCT ProgId, Program
    FROM OrganizationStructure
    ORDER BY Program
    """
    
    try:
        return q.QuerySql(sql_programs)
    except Exception as e:
        print "<p>Error retrieving programs: " + str(e) + "</p>"
        return []

def get_program_totals(data, program_id):
    """Calculate totals for a program"""
    members = 0
    previous = 0
    visitors = 0
    meetings = 0
    
    for row in data:
        if row.ProgId == program_id:
            members += row.Members
            previous += row.Previous
            visitors += row.Visitors
            meetings += row.Meetings
    
    return (members, previous, visitors, meetings)

def get_division_totals(data, division_id):
    """Calculate totals for a division"""
    members = 0
    previous = 0
    visitors = 0
    meetings = 0
    
    for row in data:
        if row.DivId == division_id:
            members += row.Members
            previous += row.Previous
            visitors += row.Visitors
            meetings += row.Meetings
    
    return (members, previous, visitors, meetings)

def parse_related_divisions(related_divs_str, current_div_id):
    """Parse the related divisions string into HTML links"""
    if not related_divs_str:
        return ""
    
    links = []
    div_parts = related_divs_str.split('|')
    
    for part in div_parts:
        if not part:
            continue
            
        try:
            div_id, div_name = part.split(':', 1)
            
            # Skip the current division
            if str(div_id) != str(current_div_id):
                links.append(
                    '<a href="#div_%s" onclick="highlightDivision(%s);">%s</a>' % 
                    (div_id, div_id, div_name)
                )
        except:
            continue
    
    if not links:
        return ""
    
    return "Also in: " + ", ".join(links)

def format_date(date_str):
    """Format a date string for display"""
    if not date_str:
        return ""
    
    # Extract just the date part if datetime
    date_parts = str(date_str).split(' ')
    return date_parts[0]

# Main execution
try:
    # Get URL parameters safely
    show_inactive = False
    if hasattr(model.Data, 'inactive'):
        inactive_value = model.Data.inactive
        if inactive_value == "yes":
            show_inactive = True
    
    selected_program = None
    if hasattr(model.Data, 'program'):
        program_value = model.Data.program
        if program_value and program_value != "0":
            selected_program = program_value
    
    created_after = None
    if hasattr(model.Data, 'created_after'):
        created_after_value = model.Data.created_after
        if created_after_value and created_after_value.strip():
            created_after = created_after_value.strip()
    
    # Get data using the optimized query that includes related divisions
    structure_data = get_structure_data_with_related(show_inactive, selected_program, created_after)
    programs_list = get_programs_for_filter()
    
    # Generate program filter options
    program_options = ""
    for program in programs_list:
        selected = ""
        if selected_program and selected_program == str(program.ProgId):
            selected = " selected"
        program_options += '<option value="%s"%s>%s</option>' % (
            program.ProgId, 
            selected,
            program.Program
        )
    
    # Generate structure rows
    structure_rows = ""
    current_program = None
    current_division = None
    
    for row in structure_data:
        # New program
        if current_program != row.ProgId:
            current_program = row.ProgId
            current_division = None
            
            # Calculate program totals
            program_members, program_previous, program_visitors, program_meetings = get_program_totals(
                structure_data, row.ProgId
            )
            
            program_row_id = "prog_" + str(row.ProgId)
            structure_rows += """
            <tr id="%s" class="program-row">
                <td>
                    <span class="toggle-icon" onclick="toggleChildren('%s', 'program')">+</span>
                    %s
                </td>
                <td>-</td>
                <td>-</td>
                <td style="text-align: center;">%s</td>
                <td style="text-align: center;">%s</td>
                <td style="text-align: center;">%s</td>
                <td style="text-align: center;">%s</td>
                <td></td>
                <td></td>
            </tr>
            """ % (
                program_row_id, 
                program_row_id, 
                row.Program,
                program_members,
                program_previous,
                program_visitors,
                program_meetings
            )
        
        # New division
        if current_division != row.DivId:
            current_division = row.DivId
            
            # Calculate division totals
            division_members, division_previous, division_visitors, division_meetings = get_division_totals(
                structure_data, row.DivId
            )
            
            division_row_id = "div_" + str(row.DivId)
            structure_rows += """
            <tr id="%s" class="division-row %s">
                <td style="padding-left: 30px;">
                    <span class="toggle-icon" onclick="toggleChildren('%s', 'division')">+</span>
                    %s
                </td>
                <td>-</td>
                <td>-</td>
                <td style="text-align: center;">%s</td>
                <td style="text-align: center;">%s</td>
                <td style="text-align: center;">%s</td>
                <td style="text-align: center;">%s</td>
                <td></td>
                <td></td>
            </tr>
            """ % (
                division_row_id, 
                "prog_" + str(row.ProgId),
                division_row_id,
                row.Division,
                division_members,
                division_previous,
                division_visitors,
                division_meetings
            )
        
        # Parse related divisions directly from the concatenated string
        related_divs_html = ""
        if hasattr(row, 'RelatedDivs'):
            related_divs_html = parse_related_divisions(row.RelatedDivs, row.DivId)
        
        # Format the created date
        created_date = format_date(row.CreatedDate)
        
        # Organization row - added OrgType column and CreatedDate column
        org_class = "inactive" if row.OrgStatus == "Inactive" else ""
        structure_rows += """
        <tr class="organization-row %s %s">
            <td style="padding-left: 60px;">
                <a href="/Org/%s" target="_blank">%s</a>
            </td>
            <td>%s</td>
            <td>%s</td>
            <td style="text-align: center;">%s</td>
            <td style="text-align: center;">%s</td>
            <td style="text-align: center;">%s</td>
            <td style="text-align: center;">%s</td>
            <td class="related-divisions">%s</td>
            <td>%s</td>
        </tr>
        """ % (
            "div_" + str(row.DivId),
            org_class,
            row.OrgId,
            row.Organization,
            row.OrgStatus,
            row.OrgType,
            row.Members,
            row.Previous,
            row.Visitors,
            row.Meetings,
            related_divs_html,
            created_date
        )
    
    # Render summary stats with filters applied
    summary_data = get_summary_stats(show_inactive, selected_program, created_after)
    
    if summary_data:
        summary_stats_html = """
        <div style="display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px;">
            <div style="flex: 1; min-width: 100px; background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; text-align: center;">
                <div style="font-size: 24px; font-weight: bold;">%s</div>
                <div style="margin-top: 5px;">Programs</div>
            </div>
            <div style="flex: 1; min-width: 100px; background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; text-align: center;">
                <div style="font-size: 24px; font-weight: bold;">%s</div>
                <div style="margin-top: 5px;">Divisions</div>
            </div>
            <div style="flex: 1; min-width: 100px; background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; text-align: center;">
                <div style="font-size: 24px; font-weight: bold;">%s</div>
                <div style="margin-top: 5px;">Total Inv</div>
            </div>
            <div style="flex: 1; min-width: 100px; background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; text-align: center;">
                <div style="font-size: 24px; font-weight: bold;">%s</div>
                <div style="margin-top: 5px;">Active Inv</div>
            </div>
            <div style="flex: 1; min-width: 100px; background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; text-align: center;">
                <div style="font-size: 24px; font-weight: bold;">%s</div>
                <div style="margin-top: 5px;">Members</div>
            </div>
            <div style="flex: 1; min-width: 100px; background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; text-align: center;">
                <div style="font-size: 24px; font-weight: bold;">%s</div>
                <div style="margin-top: 5px;">Previous</div>
            </div>
            <div style="flex: 1; min-width: 100px; background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; text-align: center;">
                <div style="font-size: 24px; font-weight: bold;">%s</div>
                <div style="margin-top: 5px;">Visitors</div>
            </div>
            <div style="flex: 1; min-width: 100px; background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; text-align: center;">
                <div style="font-size: 24px; font-weight: bold;">%s</div>
                <div style="margin-top: 5px;">Meetings</div>
            </div>
        </div>
        """ % (
            summary_data.ProgramCount,
            summary_data.DivisionCount,
            summary_data.InvCount,
            summary_data.ActiveInvs,
            summary_data.TotalMembers,
            summary_data.TotalPrevious,
            summary_data.TotalVisitors,
            summary_data.TotalMeetings
        )
    else:
        summary_stats_html = "<p>No summary data available</p>"
    
    # Output the complete HTML directly 
    print """
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 20px;
            }
            .dashboard-container {
                max-width: 1200px;
                margin: 0 auto;
            }
            .filters {
                margin-bottom: 20px;
                padding: 15px;
                background-color: #f5f5f5;
                border-radius: 5px;
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                gap: 15px;
            }
            .filters label {
                margin-right: 5px;
            }
            .filters select, .filters input[type="date"] {
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 4px;
            }
            .filters input[type="checkbox"] {
                margin-right: 5px;
            }
            .filters button {
                padding: 6px 12px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
            }
            .filters button:hover {
                background-color: #45a049;
            }
            .filter-group {
                display: flex;
                align-items: center;
                flex-wrap: wrap;
                gap: 5px;
            }
            .structure-tree {
                border: 1px solid #ddd;
                border-radius: 5px;
                overflow: auto;
            }
            table {
                width: 100%%;
                border-collapse: collapse;
            }
            th {
                background-color: #f2f2f2;
                padding: 10px;
                text-align: left;
                border-bottom: 2px solid #ddd;
            }
            td {
                padding: 8px 10px;
                border-bottom: 1px solid #ddd;
            }
            tr:hover {
                background-color: #f9f9f9;
            }
            .program-row {
                background-color: #e7f3fe;
                font-weight: bold;
                cursor: pointer;
            }
            .division-row {
                background-color: #f0f7fa;
                cursor: pointer;
                display: none;
            }
            .organization-row {
                display: none;
            }
            .organization-row.inactive {
                background-color: #ffeeee;
                color: #999;
            }
            .toggle-icon {
                display: inline-block;
                width: 20px;
                text-align: center;
                font-weight: bold;
            }
            .dashboard-header {
                margin-bottom: 20px;
                text-align: center;
            }
            .related-divisions {
                font-size: 12px;
                color: #666;
            }
            .related-divisions a {
                color: #337ab7;
                text-decoration: none;
            }
            .related-divisions a:hover {
                text-decoration: underline;
            }
            .highlight {
                animation: highlight-animation 2s ease-in-out;
            }
            @keyframes highlight-animation {
                0%% { background-color: #ffffcc; }
                100%% { background-color: transparent; }
            }
            @media print {
                .filters {
                    display: none;
                }
                .program-row, .division-row, .organization-row {
                    display: table-row !important;
                }
            }
        </style>
    </head>
    <body>
        <div class="dashboard-container">
            <div class="dashboard-header">
                <h1>Ministry Structure
                  <svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">i</text>
                  </svg>
                </h1>
            </div>
            
            <!-- Summary Statistics -->
            %s
            
            <!-- Filters -->
            <form class="filters" method="get">
                <div class="filter-group">
                    <label for="program">Program:</label>
                    <select id="program" name="program">
                        <option value="0">All Programs</option>
                        %s
                    </select>
                </div>
                
                <div class="filter-group">
                    <label for="created_after">Created After:</label>
                    <input type="date" id="created_after" name="created_after" value="%s">
                </div>
                
                <div class="filter-group">
                    <label>
                        <input type="checkbox" id="inactive" name="inactive" value="yes" %s>
                        Show Inactive Involvements
                    </label>
                </div>
                
                <button type="submit">Apply Filters</button>
            </form>
            
            <!-- Structure Tree -->
            <div class="structure-tree">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 20%%;">Name</th>
                            <th style="width: 7%%;">Status</th>
                            <th style="width: 8%%;">Type</th>
                            <th style="width: 7%%;">Members</th>
                            <th style="width: 7%%;">Previous</th>
                            <th style="width: 7%%;">Visitors</th>
                            <th style="width: 7%%;">Meetings</th>
                            <th style="width: 22%%;">Related Divisions</th>
                            <th style="width: 15%%;">Created Date</th>
                        </tr>
                    </thead>
                    <tbody id="structure-body">
                        %s
                    </tbody>
                </table>
            </div>
        </div>
        
        <script>
            // Function to toggle display of children
            function toggleChildren(elementId, type) {
                var element = document.getElementById(elementId);
                var icon = element.querySelector('.toggle-icon');
                var children = document.getElementsByClassName(elementId);
                
                var isExpanded = icon.innerHTML === "−";
                
                if (isExpanded) {
                    // Collapse
                    icon.innerHTML = "+";
                    for (var i = 0; i < children.length; i++) {
                        children[i].style.display = "none";
                        
                        // If collapsing a program, also collapse its divisions' organizations
                        if (type === 'program') {
                            var divId = children[i].id;
                            var divChildren = document.getElementsByClassName(divId);
                            var divIcon = children[i].querySelector('.toggle-icon');
                            if (divIcon) divIcon.innerHTML = "+";
                            
                            for (var j = 0; j < divChildren.length; j++) {
                                divChildren[j].style.display = "none";
                            }
                        }
                    }
                } else {
                    // Expand
                    icon.innerHTML = "−";
                    for (var i = 0; i < children.length; i++) {
                        children[i].style.display = "table-row";
                    }
                }
            }
            
            // Function to highlight a division when clicked from a related division link
            function highlightDivision(divId) {
                var divisionElement = document.getElementById('div_' + divId);
                
                // Find the program id this division belongs to
                var programClass = divisionElement.className.split(' ').find(function(cls) {
                    return cls.startsWith('prog_');
                });
                
                // Expand the program if needed
                var programElement = document.getElementById(programClass);
                var programIcon = programElement.querySelector('.toggle-icon');
                if (programIcon.innerHTML === "+") {
                    toggleChildren(programClass, 'program');
                }
                
                // Expand the division
                var divisionIcon = divisionElement.querySelector('.toggle-icon');
                if (divisionIcon.innerHTML === "+") {
                    toggleChildren('div_' + divId, 'division');
                }
                
                // Scroll to the division
                divisionElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
                
                // Add highlight effect
                divisionElement.classList.add('highlight');
                
                // Remove highlight after animation completes
                setTimeout(function() {
                    divisionElement.classList.remove('highlight');
                }, 2000);
                
                // Prevent default anchor behavior
                return false;
            }
            
            // Expand all programs by default
            window.onload = function() {
                var programRows = document.querySelectorAll('.program-row');
                for (var i = 0; i < programRows.length; i++) {
                    var programId = programRows[i].id;
                    toggleChildren(programId, 'program');
                }
            };
        </script>
    </body>
    </html>
    """ % (
        summary_stats_html,
        program_options,
        created_after or "",
        "checked" if show_inactive else "",
        structure_rows
    )

except Exception as e:
    # Print any errors
    import traceback
    print "<h2>Error</h2>"
    print "<p>An error occurred: " + str(e) + "</p>"
    print "<pre>"
    traceback.print_exc()
    print "</pre>"
