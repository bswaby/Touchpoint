ProgramID = model.Data.ProgramID
ProgramName = model.Data.ProgramName

model.Header = ProgramName + ' Balance'

sqlDueTotals = '''
SELECT 
  SUM(CASE WHEN (dbo.TransactionSummary.TranDate <= GetDate() 
             AND dbo.TransactionSummary.TranDate >= DateAdd(mm,-1, GetDate())) 
        THEN dbo.TransactionSummary.TotDue END) AS [CurrentDue],
  SUM(CASE WHEN (dbo.TransactionSummary.TranDate <= DateAdd(mm,-1, GetDate()) 
              AND dbo.TransactionSummary.TranDate >= DateAdd(mm,-3, GetDate())) 
         THEN dbo.TransactionSummary.TotDue END) AS [PastDue30-90],
  SUM(CASE WHEN (dbo.TransactionSummary.TranDate <= DateAdd(mm,-3, GetDate()) 
              AND dbo.TransactionSummary.TranDate >= DateAdd(mm,-6, GetDate())) 
         THEN dbo.TransactionSummary.TotDue END) AS [PastDue90-180] 
FROM 
    dbo.ProgDiv 
INNER JOIN 
    dbo.Program 
ON 
    ( 
        dbo.ProgDiv.ProgId = dbo.Program.Id) 
INNER JOIN 
    dbo.Division 
ON 
    ( 
        dbo.ProgDiv.DivId = dbo.Division.Id) 
INNER JOIN 
    dbo.Organizations Organizations_alias1 
ON 
    ( 
        dbo.Division.Id = Organizations_alias1.DivisionId) 
INNER JOIN 
    dbo.TransactionSummary 
ON 
    ( 
        Organizations_alias1.OrganizationId = dbo.TransactionSummary.OrganizationId) 
LEFT OUTER JOIN 
    dbo.People 
ON 
    ( 
        dbo.TransactionSummary.PeopleId = dbo.People.PeopleId) 
WHERE 
    dbo.Program.Id = ''' + ProgramID + ''' 
--AND dbo.TransactionSummary.TotDue > 0;
'''

sqlBalance = '''
SELECT 
        Organizations_alias1.OrganizationName, 
        Organizations_alias1.OrganizationId, 
        dbo.People.HomePhone, 
        dbo.People.CellPhone, 
        dbo.People.Name AS [MemberName], 
        dbo.TransactionSummary.TotDue,
        dbo.TransactionSummary.TranDate,
        FORMAT(dbo.TransactionSummary.TotPaid, 'C') AS [TotalPaid],
        dbo.People.FamilyId, 
        dbo.People.PeopleId, 
        dbo.People.PositionInFamilyId, 
        dbo.People.FirstName, 
        dbo.People.LastName, 
        dbo.People.EmailAddress, 
        dbo.People.PrimaryAddress, 
        dbo.People.PrimaryCity, 
        dbo.People.PrimaryState, 
        dbo.People.PrimaryZip, 
        dbo.People.Age, 
        dbo.TransactionSummary.RegId,
        dbo.MemberTags.Name AS [SubGroup]
    FROM 
        dbo.OrganizationMembers 
    INNER JOIN dbo.People 
      ON (dbo.OrganizationMembers.PeopleId = dbo.People.PeopleId) 
    INNER JOIN dbo.Organizations Organizations_alias1 
      ON (dbo.OrganizationMembers.OrganizationId = Organizations_alias1.OrganizationId) 
    INNER JOIN dbo.TransactionSummary 
      ON (dbo.OrganizationMembers.TranId = dbo.TransactionSummary.RegId) 
    LEFT JOIN dbo.OrgMemMemTags 
      ON (dbo.OrgMemMemTags.PeopleId = dbo.People.PeopleId) 
      AND (dbo.OrgMemMemTags.OrgId = Organizations_alias1.OrganizationId)
    LEFT JOIN dbo.MemberTags
      ON (dbo.MemberTags.Id = dbo.OrgMemMemTags.MemberTagId)
    LEFT JOIN dbo.DivOrg
      ON (dbo.DivOrg.OrgId = Organizations_alias1.OrganizationId)
    LEFT JOIN dbo.Division
      ON (dbo.Division.Id = dbo.DivOrg.DivId)
    LEFT JOIN dbo.ProgDiv
      ON (dbo.ProgDiv.DivId = dbo.Division.Id)
    INNER JOIN dbo.Program 
      ON (dbo.ProgDiv.ProgId = dbo.Program.Id) 
	--INNER JOIN dbo.Division 
	--	ON (dbo.ProgDiv.DivId = dbo.Division.Id) 
    WHERE 
    	dbo.Program.Id = ''' + ProgramID + ''' 
    	AND dbo.TransactionSummary.TotDue <> 0
        --dbo.People.PeopleId = 38975--@P1--3134;
    ORDER BY 
        dbo.TransactionSummary.TranDate Desc

'''

sqlBalance1 = '''
SELECT 
    dbo.People.FamilyId, 
    dbo.People.PeopleId, 
    dbo.People.Name AS [MemberName], 
    dbo.People.CellPhone,
    dbo.People.HomePhone,
    dbo.TransactionSummary.TranDate, 
    Organizations_alias1.OrganizationName,
    dbo.MemberTags.Name AS [SubGroup],
    dbo.TransactionSummary.TotDue
FROM 
    dbo.ProgDiv 
INNER JOIN 
    dbo.Program 
ON 
    ( 
        dbo.ProgDiv.ProgId = dbo.Program.Id) 
INNER JOIN 
    dbo.Division 
ON 
    ( 
        dbo.ProgDiv.DivId = dbo.Division.Id) 
INNER JOIN 
    dbo.Organizations Organizations_alias1 
ON 
    ( 
        dbo.Division.Id = Organizations_alias1.DivisionId) 
INNER JOIN 
    dbo.TransactionSummary 
ON 
    ( 
        Organizations_alias1.OrganizationId = dbo.TransactionSummary.OrganizationId) 
LEFT OUTER JOIN 
    dbo.People 
ON 
    ( 
        dbo.TransactionSummary.PeopleId = dbo.People.PeopleId) 
LEFT JOIN
    dbo.OrgMemMemTags
ON
    (dbo.OrgMemMemTags.PeopleId = dbo.People.PeopleId) AND (dbo.OrgMemMemTags.OrgId = Organizations_alias1.OrganizationId)
LEFT JOIN 
    dbo.MemberTags
ON
    (dbo.MemberTags.Id = dbo.OrgMemMemTags.MemberTagId)
WHERE 
    dbo.Program.Id = ''' + ProgramID + ''' 
AND dbo.TransactionSummary.TotDue > 0
Order By     dbo.People.FamilyId, 
    dbo.People.PeopleId;
'''

template = '''
<script src ="https://code.jquery.com/jquery-3.5.1.min.js"></script>                 
<script src ="https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.2/jquery-ui.min.js"> </script>  
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.0.0/jquery.min.js"></script>
<!-- jQuery Modal -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css" />

<a onclick="history.back()"><i class="fa fa-hand-o-left fa-2x" aria-hidden="true"></i></a>&nbsp;&nbsp;
<a href="https://myfbch.com/PyScript/MM-Member%20Manager?ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + '''"><i class="fa fa-home fa-2x"></i></a>
  <table>
	   {{#each duetotals}}
		  <tr><td width="200px"><h4>CurrentDue</h4></td><td align="right"><h4>{{FmtMoney CurrentDue}}</h4></td></tr>
		  <tr><td width="200px"><h4>PastDue 30-90</h4></td><td align="right"><h4>{{FmtMoney PastDue30-90}}</h4></td></tr>
		  <tr><td width="200px"><h4>PastDue 90-180</h4></td><td align="right"><h4>{{FmtMoney PastDue90-180}}</h4></td></tr>
		{{/each}}
  </table>
    <br>
    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">

      google.charts.load('current', {'packages':['table']});
      google.charts.setOnLoadCallback(drawTable);

      function drawTable() {
        var data = new google.visualization.DataTable();
        data.addColumn('number', 'FamilyId');
        data.addColumn('string', 'PeopleId');
        data.addColumn('string', 'Name');
        data.addColumn('string', 'TranDate');
        data.addColumn('string', 'Fee Group');
        data.addColumn('number', 'Total Due');
        data.addColumn('string', '');
        data.addRows([
            {{#each balancereport}}
                [
                {{FamilyId}}, 
                '<a href="https://myfbch.com/PyScript/MM-MemberDetails?p1={{PeopleId}}&FamilyId={{FamilyId}}&ProgramName=''' + ProgramName + '''&ProgramID=''' + ProgramID + '''" target="_blank">{{PeopleId}}</a>',
                '{{MemberName}}', 
                '{{TranDate}}',
                '{{OrganizationName}} ({{SubGroup}})',
                {v: {{TotDue}},  f: '{{FmtMoney TotDue}}'},
                '<a href="/Person2/{{PeopleId}}#tab-current" target="_blank"><i class="fa fa-info-circle" aria-hidden="true"></i></a>',
                ],
            {{/each}}
        ]);
        
        var table = new google.visualization.Table(document.getElementById('table_div'));
        table.draw(data, {showRowNumber: false, alternatingRowStyle: true, allowHtml: true, width: '100%', height: '100%'});
      }
    </script>
    <div id='table_div' style='width: 900px;'></div>
'''
Data.duetotals = q.QuerySql(sqlDueTotals)
Data.balancereport = q.QuerySql(sqlBalance)
NMReport = model.RenderTemplate(template)
print(NMReport)