# Division Attendance 
# This script creates a dashboard showing attendance across divisions by Sunday
# 
# --Upload Instructions Start--
# To upload code to Touchpoint, use the following steps.
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File
# 3. Name the Python and paste all this code
# 4. Test and optionally add to menu
# --Upload Instructions End--

model.Header = "Division Attendance"

# Define a safer function to get parameters with defaults
def safe_get_param(param_name, default_value):
    if hasattr(model.Data, param_name):
        value = getattr(model.Data, param_name)
        if value:
            return value
    return default_value

# Get URL parameters
begindate = safe_get_param('begindate', None)
enddate = safe_get_param('enddate', None)
divs = safe_get_param('divs', None)

# Format begindate and enddate for SQL
if begindate:
    begindate_sql = "'" + begindate + "'"
else:
    begindate_sql = "DATEADD(MONTH, -6, GETDATE())"
    
if enddate:
    enddate_sql = "'" + enddate + "'"
else:
    enddate_sql = "GETDATE()"
    
# Format divs for SQL - default to a sample if not provided
if not divs:
    divs = "7,8,9,10"  # Replace with default division IDs

# Simpler SQL that doesn't use a temporary table for division IDs
sql = """
DECLARE @begindate DATE = """ + begindate_sql + """;
DECLARE @enddate DATE = """ + enddate_sql + """;

-- Find the Sunday for each meeting date and sum attendance by division
SELECT 
    CONVERT(VARCHAR, 
        DATEADD(DAY, 
            -(DATEPART(WEEKDAY, m.MeetingDate) - 1), 
            m.MeetingDate), 
        101) AS MeetingDate,
    d.Name AS DivisionName,
    SUM(COALESCE(m.MaxCount, 0)) AS Attendance
FROM dbo.Meetings m
JOIN dbo.Organizations o ON o.OrganizationId = m.OrganizationId
JOIN dbo.DivOrg dd ON dd.OrgId = o.OrganizationId
JOIN dbo.Division d ON d.Id = dd.DivId
WHERE dd.DivId IN (""" + divs + """)
AND m.MeetingDate BETWEEN @begindate AND @enddate
GROUP BY 
    DATEADD(DAY, 
        -(DATEPART(WEEKDAY, m.MeetingDate) - 1), 
        m.MeetingDate),
    d.Name
ORDER BY 
    DATEADD(DAY, 
        -(DATEPART(WEEKDAY, m.MeetingDate) - 1), 
        m.MeetingDate);
"""

try:
    # Execute the SQL query
    results = q.QuerySql(sql)
    
    # Debug info
    #print "<p>SQL query executed. Results found: " + str(len(results)) + "</p>"
    
    # Initialize variables
    divisions = []
    sundays = []
    attendance_data = []
    
    # Process the SQL results to extract data - with defensive programming
    if results is not None:
        for row in results:
            if hasattr(row, 'MeetingDate') and hasattr(row, 'DivisionName') and hasattr(row, 'Attendance'):
                sunday = row.MeetingDate
                division = row.DivisionName
                attendance = row.Attendance if row.Attendance is not None else 0
                
                # Add to divisions list if not already there
                if division not in divisions:
                    divisions.append(division)
                
                # Add to sundays list if not already there
                if sunday not in sundays:
                    sundays.append(sunday)
                
                # Add the data point
                attendance_data.append({
                    'Sunday': sunday,
                    'Division': division,
                    'Attendance': attendance
                })
    
    # Sort sundays chronologically
    sundays.sort()
    
    html = """<style>
        .date-picker {{ margin-bottom: 20px; }}
        .date-picker label {{ margin-right: 10px; }}
        .date-picker input, .date-picker select {{ margin-right: 20px; }}
        .chart-container {{ height: 500px; width: 100%; }}
    </style>
    
    <div class='date-picker'>
        <form method='get'>
            <label for='begindate'>Start Date:</label>
            <input type='date' id='begindate' name='begindate' value='{}'>
            <label for='enddate'>End Date:</label>
            <input type='date' id='enddate' name='enddate' value='{}'>
            <label for='divs'>Division IDs:</label>
            <input type='text' id='divs' name='divs' value='{}'>
            <button type='submit'>Apply Filter</button>
        </form>
    </div>
    
    <div class='chart-container'>
        <script type='text/javascript' src='https://www.gstatic.com/charts/loader.js'></script>
        <script type='text/javascript'>
            google.charts.load('current', {{'packages':['corechart']}});
            google.charts.setOnLoadCallback(drawChart);
            function drawChart() {{
                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Sunday');
    """.format(begindate or "", enddate or "", divs)


    
    # Add columns for divisions
    for div in divisions:
        html += "data.addColumn('number', '" + div + "');"
    
    # Create data rows
    html += "data.addRows(["
    
    # If we have data, populate it
    if len(sundays) > 0 and len(divisions) > 0:
        for sunday in sundays:
            row_str = "['" + sunday + "'"
            
            for div in divisions:
                # Find attendance for this division on this Sunday
                attendance = 0
                for data_point in attendance_data:
                    if data_point['Sunday'] == sunday and data_point['Division'] == div:
                        attendance = data_point['Attendance']
                        break
                
                row_str += ", " + str(attendance)
            
            row_str += "],"
            html += row_str
    else:
        # Add a dummy row if no data
        html += "['No Data', 0],"
    
    # Finish the chart code
    html += "]);"
    html += "var options = {"
    html += "title: 'Program Attendance by Sunday',"
    html += "curveType: 'function',"
    html += "legend: { position: 'bottom' },"
    html += "chartArea: {width: '80%', height: '70%'}"
    html += "};"
    html += "var chart = new google.visualization.LineChart(document.getElementById('chart_div'));"
    html += "chart.draw(data, options);"
    html += "}"
    html += "</script>"
    html += "<div id='chart_div' style='width: 100%; height: 500px'></div>"
    html += "</div>"
    
    # Add data table
    html += "<h3>Data Table</h3>"
    html += "<table class='table table-striped' style='width: 100%;'>"
    html += "<thead><tr><th>Sunday</th>"
    
    for div in divisions:
        html += "<th>" + div + "</th>"
    
    html += "</tr></thead><tbody>"
    
    # If we have data, populate the table
    if len(sundays) > 0:
        for sunday in sundays:
            html += "<tr><td>" + sunday + "</td>"
            
            for div in divisions:
                # Find attendance for this division on this Sunday
                attendance = 0
                for data_point in attendance_data:
                    if data_point['Sunday'] == sunday and data_point['Division'] == div:
                        attendance = data_point['Attendance']
                        break
                
                html += "<td>" + str(attendance) + "</td>"
            
            html += "</tr>"
    else:
        html += "<tr><td colspan='" + str(len(divisions) + 1) + "'>No data available</td></tr>"
    
    html += "</tbody></table>"
    
    # Print the HTML
    print html

except Exception as e:
    # Print any errors
    import traceback
    print "<h2>Error</h2>"
    print "<p>An error occurred: " + str(e) + "</p>"
    print "<pre>"
    traceback.print_exc()
    print "</pre>"
