#####################################################################
# Registration Question Export with Question Consolidation
#####################################################################
# This script creates a CSV export of registration data that:
# 1. Consolidates identical questions from different orgs into single columns
# 2. Groups all answers for a person into a single row
# 3. Formats the data for easy import into spreadsheet applications
# 4. Supports multiple organizations with parameter filtering
#
#####################################################################
# Upload Instructions
#####################################################################
# To upload code to Touchpoint, use the following steps:
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File
# 3. Name the Python (suggested: RegQuestionExport) and paste all this code
# 4. Test and optionally add to menu
#####################################################################

model.Header = "Registration Question Export"

# Get organization IDs from parameter (p1) or set defaults
org_ids = model.Data.p1 if hasattr(model.Data, 'p1') and model.Data.p1 else ""

# Display error message
def show_error(error_message):
    print """
    <div style="color: red; background-color: #ffeeee; padding: 10px; border-radius: 5px; margin: 20px 0; border: 1px solid #ffcccc;">
        <h3>Error</h3>
        <p>{0}</p>
    </div>
    """.format(error_message)

# Display loading indicator
def show_loading():
    print """
    <div id="loading" style="display: block; text-align: center; padding: 20px;">
        <p>Processing data, please wait...</p>
        <div style="width: 50px; height: 50px; border: 8px solid #f3f3f3; border-top: 8px solid #3498db; border-radius: 50%; animation: spin 2s linear infinite; margin: 0 auto;"></div>
    </div>
    <style>
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .export-controls {
            background-color: #f5f5f5;
            padding: 10px;
            margin-bottom: 20px;
            border-radius: 5px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .export-controls button {
            padding: 8px 16px;
            background-color: #4CAF50;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        }
        .export-controls a {
            padding: 8px 16px;
            background-color: #f1f1f1;
            color: #333;
            text-decoration: none;
            border-radius: 4px;
            font-size: 14px;
        }
        .question-reference {
            background-color: #f9f9f9;
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
            background-color: #f0f0f0;
            padding: 8px;
            text-align: left;
        }
        .question-reference td {
            padding: 8px;
            border-top: 1px solid #ddd;
        }
        .data-table {
            overflow-x: auto;
            margin-top: 20px;
        }
        .data-table table {
            border-collapse: collapse;
            width: 100%;
            font-size: 13px;
        }
        .data-table th {
            background-color: #f0f0f0;
            padding: 8px;
            text-align: left;
            position: sticky;
            top: 0;
            border: 1px solid #ddd;
        }
        .data-table td {
            padding: 8px;
            border: 1px solid #ddd;
        }
        .data-table tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        /* Add debug borders to help spot issues */
        .debug-border {
            border: 1px solid red !important;
        }
    </style>
    <script>
        window.onload = function() {
            document.getElementById('loading').style.display = 'none';
            document.getElementById('results-container').style.display = 'block';
        };
    </script>
    """

# Get organization name from ID
def get_org_name(org_id):
    try:
        sql = "SELECT OrganizationName FROM Organizations WHERE OrganizationId = " + str(org_id)
        result = q.QuerySqlTop1(sql)
        if result and hasattr(result, 'OrganizationName'):
            return result.OrganizationName
        return "Unknown Organization"
    except:
        return "Organization #" + str(org_id)

# Get organizations with registration questions
def get_registration_orgs():
    try:
        sql = """
        SELECT DISTINCT o.OrganizationId, o.OrganizationName
        FROM Organizations o
        JOIN RegQuestion rq ON o.OrganizationId = rq.OrganizationId
        WHERE o.OrganizationStatusId = 30 
        ORDER BY o.OrganizationName
        """
        return q.QuerySql(sql)
    except:
        return []

# Display organization selection form
if not org_ids:
    print """
    <h2>Select Organizations for Question Export
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
        <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">si</text>
      </svg>
    </h2>
    <form method="get" action="">
        <div style="margin-bottom: 20px;">
            <label for="org_ids">Enter Organization IDs (comma-separated):</label>
            <input type="text" id="org_ids" name="p1" style="width: 300px;">
        </div>
        <p>- OR -</p>
        <div style="margin-bottom: 20px;">
            <label>Select from organizations with registration questions:</label>
            <div style="max-height: 300px; overflow-y: auto; border: 1px solid #ccc; padding: 10px; margin-top: 5px;">
    """
    
    orgs = get_registration_orgs()
    if orgs:
        for org in orgs:
            print """
            <div style="margin-bottom: 5px;">
                <a href="?p1={0}" style="text-decoration: none;">
                    <button type="button" style="width: 100%; text-align: left; padding: 5px;">
                        {1} (ID: {0})
                    </button>
                </a>
            </div>
            """.format(org.OrganizationId, org.OrganizationName)
    else:
        print "<p>No organizations with registration questions found.</p>"
    
    print """
            </div>
        </div>
        <div>
            <button type="submit">Generate Export</button>
        </div>
    </form>
    """
# Process data when org_ids are provided
else:
    try:
        show_loading()
        print '<div id="results-container" style="display: none;">'
        
        # Clean up org_ids input
        org_ids = org_ids.replace(" ", "")
        
        # Display header with selected organizations
        print """<h2>Registration Question Export
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
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">si</text>
                  </svg>
        </h2>"""
        print '<h3>Selected Organizations:</h3>'
        print '<ul>'
        
        org_id_list = org_ids.split(",")
        orgs_info = []
        for org_id in org_id_list:
            if org_id and org_id.isdigit():
                org_name = get_org_name(org_id)
                print '<li>{0} (ID: {1})</li>'.format(org_name, org_id)
        print '</ul>'
        
        # Get questions consolidated by label - FIXED to avoid STRING_AGG with UniqueIdentifiers
        # Use explicit ID numbering starting at 1
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
        
        # Check if we found any questions
        if not questions:
            show_error("No registration questions found for the selected organizations.")
            print '</div>'
        else:
            # Print all questions for debugging
            print "<!-- Debug: Question Data -->"
            print "<!-- Total questions found: {0} -->".format(len(questions))
            for q_idx, q_item in enumerate(questions):
                print "<!-- Q{0}: {1} -->".format(q_idx + 1, q_item.Label)

            # Define mapping for SQL column names
            question_columns = []
            question_headers = []
            column_map = {}
            
            # Ensure we start the numbering at Q1 (not Q3)
            for q_idx, q_item in enumerate(questions):
                column_name = "Q{0}".format(q_idx + 1)  # Start at Q1
                question_columns.append(column_name)
                question_headers.append(q_item.Label)
                column_map[q_item.Label] = column_name
            
            # First, display the export controls and question reference at the top
            print """
            <div class="export-controls">
                <button onclick="exportToCSV();">
                    <i class="fa fa-download"></i> Export to CSV
                </button>
                <a href="?">Select Different Organizations</a>
            </div>
            
            <div class="question-reference">
                <h3>Question Reference</h3>
                <p>The following questions have been consolidated across organizations:</p>
                <table>
                    <tr>
                        <th>Column</th>
                        <th>Question</th>
                        <th>Used in # Orgs</th>
                    </tr>
            """
            
            for q_idx, q_item in enumerate(questions):
                column_name = "Q{0}".format(q_idx + 1)  # Start at Q1
                print '<tr><td>{0}</td><td>{1}</td><td style="text-align: center;">{2}</td></tr>'.format(
                    column_name, q_item.Label, q_item.OrgCount)
            
            print """
                </table>
            </div>
            """
            
            # Get registration data
            data_sql = """
            SELECT 
                r.OrganizationId,
                o.OrganizationName,
                rp.PeopleId,
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
            LEFT JOIN RegAnswer ra ON rp.RegPeopleId = ra.RegPeopleId
            LEFT JOIN RegQuestion rq ON ra.RegQuestionId = rq.RegQuestionId
            WHERE r.OrganizationId IN ({0})
            AND rq.Label IS NOT NULL
            """.format(org_ids)
            
            reg_data = q.QuerySql(data_sql)
            
            # Process the data into a person-centric dictionary
            person_data = {}
            for record in reg_data:
                person_key = str(record.PeopleId)
                
                if person_key not in person_data:
                    person_data[person_key] = {
                        'OrganizationId': record.OrganizationId,
                        'OrganizationName': record.OrganizationName,
                        'PeopleId': record.PeopleId,
                        'FirstName': record.FirstName,
                        'LastName': record.LastName,
                        'EmailAddress': record.EmailAddress,
                        'Phone': record.Phone
                    }
                    # Initialize all question columns to empty string
                    for col in question_columns:
                        person_data[person_key][col] = ''
                
                # Map the question label to the standardized column and store the answer
                if record.QuestionLabel in column_map:
                    column = column_map[record.QuestionLabel]
                    answer_value = record.AnswerValue
                    if answer_value is not None:
                        person_data[person_key][column] = answer_value
            
            # Generate the CSV data
            csv_header = ['OrganizationId', 'OrganizationName', 'PeopleId', 'FirstName', 'LastName', 'EmailAddress', 'Phone'] + question_columns
            csv_rows = []
            
            # Add all person records to the CSV data
            for person_key, data in person_data.items():
                row = []
                for field in csv_header:
                    row.append(str(data.get(field, '')))
                csv_rows.append(row)
            
            # Sort rows by LastName, FirstName
            csv_rows.sort(key=lambda x: (x[4], x[3]))  # 4 is LastName, 3 is FirstName
            
            # Generate HTML table
            print '<div class="data-table">'
            print '<table>'
            
            # Table header with friendly names for standard columns
            print '<tr>'
            print '<th>Org ID</th><th>Organization</th><th>ID</th><th>First Name</th><th>Last Name</th><th>Email</th><th>Phone</th>'
            
            # Add simple question headers - just the question number
            for q_idx in range(len(question_headers)):
                column_name = "Q{0}".format(q_idx + 1)  # Start at Q1
                print '<th>{0}</th>'.format(column_name)
            print '</tr>'
            
            # Output data rows
            for row in csv_rows:
                print '<tr>'
                for i, cell in enumerate(row):
                    cell_value = cell if cell else ''
                    print '<td>{0}</td>'.format(cell_value)
                print '</tr>'
            
            print '</table>'
            print '</div>'
            
            # Add the CSV export JavaScript function - Made more reliable
            print """
            <script>
                function exportToCSV() {
                    // Format timestamp for filename
                    var now = new Date();
                    var timestamp = now.getFullYear() + '-' +
                                   ('0' + (now.getMonth() + 1)).slice(-2) + '-' +
                                   ('0' + now.getDate()).slice(-2);
                    
                    // Create CSV content
                    var csvRows = [];
                    
                    // Get table headers
                    var headers = [];
                    var headerCells = document.querySelectorAll('.data-table table tr:first-child th');
                    headerCells.forEach(function(cell) {
                        headers.push(cell.textContent.trim());
                    });
                    csvRows.push('"' + headers.join('","') + '"');
                    
                    // Get table data
                    var rows = document.querySelectorAll('.data-table table tr:not(:first-child)');
                    rows.forEach(function(row) {
                        var rowData = [];
                        row.querySelectorAll('td').forEach(function(cell) {
                            // Properly escape quotes in cell content
                            var content = cell.textContent.trim().replace(/"/g, '""');
                            rowData.push('"' + content + '"');
                        });
                        csvRows.push(rowData.join(','));
                    });
                    
                    // Combine rows with line breaks
                    var csvString = csvRows.join('\\r\\n');
                    
                    // Create download link
                    var blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
                    var link = document.createElement('a');
                    var url = URL.createObjectURL(blob);
                    
                    // Set link properties
                    link.setAttribute('href', url);
                    link.setAttribute('download', 'RegQuestionExport_' + timestamp + '.csv');
                    link.style.visibility = 'hidden';
                    
                    // Add to document, click and remove
                    document.body.appendChild(link);
                    link.click();
                    
                    // Clean up
                    setTimeout(function() {
                        document.body.removeChild(link);
                        window.URL.revokeObjectURL(url);
                    }, 100);
                    
                    return false;
                }
            </script>
            """
            
            print '</div>'  # Close results container
        
    except Exception as e:
        import traceback
        error_message = "An error occurred: " + str(e)
        show_error(error_message)
        print "<pre>"
        traceback.print_exc()
        print "</pre>"
